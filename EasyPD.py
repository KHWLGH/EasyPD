#!/usr/bin/env python3
# -*- coding: utf-8 -*-


import sys
import csv
import time
import textwrap
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
from multiprocessing import Process, Queue, Event, Value, freeze_support
import queue
import os
try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QTextEdit,
    QTableWidget,
    QTableWidgetItem,
    QAbstractItemView,
    QHeaderView,
    QFileDialog,
    QMessageBox,
    QSplitter,
    QCheckBox,
    QComboBox,
    QLineEdit,
    QDialogButtonBox,
    QDialog,
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFontMetrics, QPalette, QColor, QPainter, QFont
from PySide6.QtCharts import QChart, QChartView, QLineSeries, QValueAxis

from device_comm import (
    HID_AVAILABLE,
    K2_TARGET_VID,
    K2_TARGET_PID,
    data_collection_worker,
    enumerate_devices,
)
from pd_decoder import (
    CableDataParser,
    PDParser,
    is_pdo_packet,
    is_rdo_packet,
)
from i18n import (
    LANG_STRINGS,
    get_text,
    translate_cable_field,
    translate_cable_value,
)

STATUS_COLORS = {
    "status_disconnected": "#bbbbbb",
    "status_connected": "#3fc1c9",
    "status_collecting": "#4fd3c4",
    "status_paused": "#ffa931",
}

class PDViewerWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(1280, 720)

        self.current_language = "zh"
        self.current_status_key = "status_disconnected"

        self.device_open = False
        self.is_paused = True
        self.log_records: List[Dict[str, Any]] = []
        self.log_index = 0
        self.start_time: Optional[float] = None
        self.cable_info_rows: List[Tuple[str, Any]] = []

        self.summary_wrap_chars = 30
        self.summary_max_lines = 4
        self.summary_char_px = 8
        self.summary_line_height = 18
        self.summary_base_padding = 6
        self.summary_metrics: Optional[QFontMetrics] = None
        self.current_max_lines = 1
        self._updating_summary = False

        self.device_candidates: List[Dict[str, Any]] = []
        self.selected_device_value: Any = None

        self.collection_process: Optional[Process] = None
        self.data_queue: Optional[Queue] = None
        self.stop_event = None
        self.pause_flag = None

        # 测量数据记录
        self.measurement_records: List[Dict[str, Any]] = []
        self.measurement_index = 0

        # 滚动记录限制（UI显示的记录数，保证性能）
        self.ui_log_limit = 5000  # UI最多显示5000条PD记录
        self.ui_measurement_limit = 10000  # UI最多显示10000条测量记录
        self.raw_log_records: List[Dict[str, Any]] = []  # 原始完整PD记录（用于导出）
        self.raw_measurement_records: List[Dict[str, Any]] = []  # 原始完整测量记录

        # 自动暂停阈值设置
        self.auto_pause_threshold_enabled = False
        self.voltage_threshold = 0.0
        self.current_threshold = 0.0
        self.auto_pause_metric = "voltage"  # voltage | current
        self.auto_pause_delay_seconds = 0.0

        self._last_voltage: Optional[float] = None
        self._last_current: Optional[float] = None

        self._auto_pause_pending_metric: Optional[str] = None
        self._auto_pause_pending_threshold: Optional[float] = None

        self.auto_pause_delay_timer = QTimer(self)
        self.auto_pause_delay_timer.setSingleShot(True)
        self.auto_pause_delay_timer.timeout.connect(self._execute_pending_auto_pause)

        # 批量UI更新相关
        self.pending_records: List[Dict[str, Any]] = []  # 待添加的PD记录
        self.pending_measurements: List[Dict[str, Any]] = []  # 待添加的测量记录
        self.ui_update_timer = QTimer(self)
        self.ui_update_timer.timeout.connect(self._batch_update_ui)
        self.ui_update_timer.setInterval(200)  # 每200ms批量更新一次UI

        # 内存监控相关
        self.memory_check_timer = QTimer(self)
        self.memory_check_timer.timeout.connect(self._check_memory_usage)
        self.memory_check_timer.setInterval(5000)  # 每5秒检查一次
        self.memory_warning_shown = False
        self.memory_threshold_mb = 500  # 500MB警告阈值

        self._build_ui()
        self._populate_device_list()
        self._apply_language()

        self.poll_timer = QTimer(self)
        self.poll_timer.timeout.connect(self._poll_queue)
        self.poll_timer.start(50)

        # 启动批量UI更新定时器
        self.ui_update_timer.start()

        # 启动内存监控定时器
        self.memory_check_timer.start()

    def _check_memory_usage(self):
        """检查内存使用，超过阈值时发出警告"""
        try:
            if not PSUTIL_AVAILABLE:
                return

            process = psutil.Process(os.getpid())
            mem_info = process.memory_info()
            mem_mb = mem_info.rss / 1024 / 1024  # 转换为MB

            # 更新状态栏显示内存使用
            if hasattr(self, 'statusBar'):
                # 查找或创建内存标签
                status_bar = self.statusBar()
                if not hasattr(self, '_memory_label'):
                    # 添加分隔符
                    status_bar.addPermanentWidget(QLabel(" | "))
                    self._memory_label = QLabel(f"RAM: {mem_mb:.0f}MB")
                    status_bar.addPermanentWidget(self._memory_label)
                else:
                    self._memory_label.setText(f"RAM: {mem_mb:.0f}MB")

            # 检查是否超过阈值
            if mem_mb > self.memory_threshold_mb and not self.memory_warning_shown:
                self.memory_warning_shown = True
                record_count = len(self.raw_log_records) + len(self.raw_measurement_records)
                QMessageBox.warning(self, self._text("dialog_warning_title"),
                    f"内存使用警告 (已使用 {mem_mb:.0f}MB)。当前已记录 {record_count} 条数据。\n\n"
                    "建议：\n"
                    "1. 导出数据到CSV文件\n"
                    "2. 清空当前记录\n"
                    "3. 继续使用（可能影响性能）")

        except Exception:
            # 如果内存监控失败，忽略错误
            pass

    def _text(self, key: str, **kwargs) -> str:
        return get_text(self.current_language, key, **kwargs)

    def _apply_language(self) -> None:
        self.setWindowTitle(self._text("app_title"))
        if hasattr(self, "device_label"):
            self.device_label.setText(self._text("label_device"))
        if hasattr(self, "refresh_devices_btn"):
            self.refresh_devices_btn.setText(self._text("btn_refresh_devices"))
        if hasattr(self, "clear_btn"):
            self.clear_btn.setText(self._text("btn_clear"))
        if hasattr(self, "export_btn"):
            self.export_btn.setText(self._text("btn_export"))
        if hasattr(self, "import_btn"):
            self.import_btn.setText(self._text("btn_import"))
        if hasattr(self, "auto_scroll_checkbox"):
            self.auto_scroll_checkbox.setText(self._text("checkbox_auto_scroll"))
        if hasattr(self, "cable_label"):
            self.cable_label.setText(self._text("cable_title"))
        if hasattr(self, "pdo_label"):
            self.pdo_label.setText(self._text("pdo_title"))
        if hasattr(self, "detail_label"):
            self.detail_label.setText(self._text("detail_title"))
        if hasattr(self, "records_label"):
            self.records_label.setText(self._text("records_title"))
        self._update_measurement_count_label()  # 更新测量数据计数显示

        self._update_connect_button_text()
        self._update_start_button_text()
        self._set_status_message(self.current_status_key)
        self._update_count_label()
        self._update_table_headers()
        self._update_cable_info(self.cable_info_rows)
        self._refresh_device_selector_labels()

        if hasattr(self, "data_visualization_checkbox"):
            self.data_visualization_checkbox.setText(self._text("checkbox_data_visualization"))

        if hasattr(self, "auto_pause_btn"):
            self.auto_pause_btn.setText(self._text("btn_auto_pause_settings"))

        # 更新图表翻译
        self._update_chart_translations()
        self._update_auto_pause_status_display()

        if hasattr(self, "auto_pause_status_label"):
            self.auto_pause_status_label.setToolTip(self._text("auto_pause_status_tooltip"))

        if hasattr(self, "language_checkbox"):
            self.language_checkbox.blockSignals(True)
            self.language_checkbox.setChecked(self.current_language == "en")
            self.language_checkbox.blockSignals(False)

    def _update_chart_translations(self) -> None:
        """更新图表的翻译文本"""
        if not hasattr(self, "unified_chart"):
            return

        # 更新图表标题
        self.unified_chart.setTitle(self._text("chart_title"))

        # 更新X轴和Y轴标签
        if hasattr(self, "unified_axis_x"):
            self.unified_axis_x.setTitleText(self._text("chart_axis_x"))
        if hasattr(self, "unified_axis_y"):
            self.unified_axis_y.setTitleText(self._text("chart_axis_y"))

        # 更新曲线名称（图例）
        if hasattr(self, "voltage_series"):
            self.voltage_series.setName(self._text("chart_voltage"))
        if hasattr(self, "current_series"):
            self.current_series.setName(self._text("chart_current"))
        if hasattr(self, "power_series"):
            self.power_series.setName(self._text("chart_power"))

    def _update_connect_button_text(self) -> None:
        if not hasattr(self, "connect_btn"):
            return
        text = self._text("btn_disconnect") if self.device_open else self._text("btn_connect")
        self.connect_btn.setText(text)

    def _update_start_button_text(self) -> None:
        if not hasattr(self, "start_btn"):
            return
        if not self.device_open:
            text = self._text("btn_start")
        elif self.is_paused:
            if self.start_time is None:
                text = self._text("btn_start")
            else:
                text = self._text("btn_resume")
        else:
            text = self._text("btn_pause")
        self.start_btn.setText(text)

    def _refresh_device_selector_labels(self) -> None:
        if not hasattr(self, "device_selector"):
            return
        self.device_selector.blockSignals(True)
        for idx, entry in enumerate(self.device_candidates):
            value = entry.get("value")
            if value is None:
                label = self._text("device_auto")
                tooltip = self._text("tooltip_default_vidpid", vid=K2_TARGET_VID, pid=K2_TARGET_PID)
            else:
                label = entry.get("label")
                if not label or label.startswith("设备") or label.startswith("Device"):
                    label = self._text("device_fallback_label", index=max(1, idx))
                vid = value.get("vid", K2_TARGET_VID)
                pid = value.get("pid", K2_TARGET_PID)
                tooltip_lines = [self._text("tooltip_vid_pid", vid=vid, pid=pid)]
                manufacturer = value.get("manufacturer")
                if manufacturer:
                    tooltip_lines.append(self._text("tooltip_manufacturer", manufacturer=manufacturer))
                product = value.get("product")
                if product:
                    tooltip_lines.append(self._text("tooltip_product", product=product))
                serial = value.get("serial")
                if serial:
                    tooltip_lines.append(self._text("tooltip_serial", serial=serial))
                path = value.get("path")
                if isinstance(path, (bytes, bytearray)):
                    try:
                        path_display = path.decode("utf-8")
                    except Exception:
                        path_display = str(path)
                else:
                    path_display = str(path) if path is not None else ""
                if path_display:
                    tooltip_lines.append(self._text("tooltip_path", path=path_display))
                tooltip = "\n".join(line for line in tooltip_lines if line)
            entry["label"] = label
            entry["tooltip"] = tooltip
            if idx < self.device_selector.count():
                self.device_selector.setItemText(idx, label)
                if tooltip:
                    self.device_selector.setItemData(idx, tooltip, Qt.ToolTipRole)
        self.device_selector.blockSignals(False)

    def _set_status_message(self, key: str, update_style: bool = True) -> None:
        self.current_status_key = key
        if hasattr(self, "status_label"):
            self.status_label.setText(self._text(key))
            if update_style:
                color = STATUS_COLORS.get(key, STATUS_COLORS["status_disconnected"])
                self.status_label.setStyleSheet(f"color: {color};")

    def _update_count_label(self) -> None:
        if hasattr(self, "count_label"):
            self.count_label.setText(self._text("records_count", count=len(self.log_records)))

    def _update_measurement_count_label(self) -> None:
        """更新测量数据计数显示"""
        if hasattr(self, "measurement_count_label"):
            self.measurement_count_label.setText(self._text("measurement_count", count=len(self.measurement_records)))

    def _translate_cable_field(self, key: str) -> str:
        return translate_cable_field(self.current_language, key)

    def _translate_cable_value(self, value: Any) -> str:
        return translate_cable_value(self.current_language, value)

    def _update_table_headers(self) -> None:
        if hasattr(self, "table"):
            self.table.setHorizontalHeaderLabels([
                self._text("table_column_index"),
                self._text("table_column_relative_time"),
                self._text("table_column_type"),
                self._text("table_column_summary"),
            ])
        if hasattr(self, "current_pdo_table"):
            self.current_pdo_table.setHorizontalHeaderLabels([
                self._text("pdo_column_index"),
                self._text("pdo_column_description"),
            ])
        if hasattr(self, "cable_info_table"):
            self.cable_info_table.setHorizontalHeaderLabels([
                self._text("cable_column_field"),
                self._text("cable_column_value"),
            ])

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(12, 12, 12, 12)
        root_layout.setSpacing(8)

        control_layout = QHBoxLayout()
        control_layout.setSpacing(6)

        self.device_label = QLabel()
        control_layout.addWidget(self.device_label)

        self.device_selector = QComboBox()
        self.device_selector.setMinimumWidth(160)
        self.device_selector.currentIndexChanged.connect(self._on_device_selected)
        control_layout.addWidget(self.device_selector)

        self.refresh_devices_btn = QPushButton()
        self.refresh_devices_btn.clicked.connect(self._populate_device_list)
        control_layout.addWidget(self.refresh_devices_btn)

        self.connect_btn = QPushButton()
        self.connect_btn.clicked.connect(self._toggle_connection)
        control_layout.addWidget(self.connect_btn)

        self.start_btn = QPushButton()
        self.start_btn.setEnabled(False)
        self.start_btn.clicked.connect(self._toggle_collection)
        control_layout.addWidget(self.start_btn)

        self.clear_btn = QPushButton()
        self.clear_btn.clicked.connect(self._confirm_manual_clear)
        control_layout.addWidget(self.clear_btn)

        self.export_btn = QPushButton()
        self.export_btn.clicked.connect(self._export_csv)
        control_layout.addWidget(self.export_btn)

        self.import_btn = QPushButton()
        self.import_btn.clicked.connect(self._import_csv)
        control_layout.addWidget(self.import_btn)

        self.auto_scroll_checkbox = QCheckBox()
        self.auto_scroll_checkbox.setChecked(True)
        control_layout.addWidget(self.auto_scroll_checkbox)

        self.data_visualization_checkbox = QCheckBox(self._text("checkbox_data_visualization"))
        self.data_visualization_checkbox.setTristate(False)
        self.data_visualization_checkbox.toggled.connect(self._on_data_visualization_toggled)
        control_layout.addWidget(self.data_visualization_checkbox)

        self.auto_pause_btn = QPushButton(self._text("btn_auto_pause_settings"))
        self.auto_pause_btn.clicked.connect(self._show_auto_pause_settings)
        control_layout.addWidget(self.auto_pause_btn)

        self.language_checkbox = QCheckBox("English")
        self.language_checkbox.setTristate(False)
        self.language_checkbox.toggled.connect(self._on_language_checkbox_toggled)
        control_layout.addWidget(self.language_checkbox)

        control_layout.addStretch(1)

        root_layout.addLayout(control_layout)

        self.splitter = QSplitter(Qt.Horizontal)
        root_layout.addWidget(self.splitter, 1)

        # 创建PD记录列表容器（最左侧）
        records_container = QWidget()
        records_layout = QVBoxLayout(records_container)
        records_layout.setContentsMargins(0, 0, 0, 0)
        records_layout.setSpacing(4)

        self.records_label = QLabel()
        self.records_label.setStyleSheet("font-weight: bold;")
        records_layout.addWidget(self.records_label)

        self.table = QTableWidget(0, 4)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setWordWrap(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        header = self.table.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignCenter)
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        # 确保序号列有足够的宽度
        self.table.setColumnWidth(0, 60)  # 设置序号列最小宽度为60像素
        header.sectionResized.connect(self._on_summary_column_resized)
        self.table.itemSelectionChanged.connect(self._on_selection_changed)
        records_layout.addWidget(self.table, 1)

        self.splitter.addWidget(records_container)

        # 添加详细信息容器（在PD记录列表右侧）
        detail_container = QWidget()
        detail_layout = QVBoxLayout(detail_container)
        detail_layout.setContentsMargins(8, 0, 0, 0)
        detail_layout.setSpacing(4)

        self.detail_label = QLabel()
        self.detail_label.setStyleSheet("font-weight: bold;")
        detail_layout.addWidget(self.detail_label)

        self.detail_text = QTextEdit()
        self.detail_text.setReadOnly(True)
        detail_layout.addWidget(self.detail_text, 1)

        self.splitter.addWidget(detail_container)

        # 创建PDO容器（中间第三个）
        self.pdo_container = QWidget()
        pdo_layout = QVBoxLayout(self.pdo_container)
        pdo_layout.setContentsMargins(0, 0, 0, 0)
        pdo_layout.setSpacing(4)

        self.pdo_label = QLabel()
        self.pdo_label.setStyleSheet("font-weight: bold;")
        pdo_layout.addWidget(self.pdo_label)

        self.current_pdo_table = QTableWidget(0, 2)
        self.current_pdo_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.current_pdo_table.setSelectionMode(QAbstractItemView.NoSelection)
        self.current_pdo_table.verticalHeader().setVisible(False)
        self.current_pdo_table.horizontalHeader().setStretchLastSection(True)
        self.current_pdo_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.current_pdo_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        pdo_layout.addWidget(self.current_pdo_table, 1)

        self.splitter.addWidget(self.pdo_container)

        # 创建线缆信息容器（最右侧第四个）
        self.cable_container = QWidget()
        cable_layout = QVBoxLayout(self.cable_container)
        cable_layout.setContentsMargins(0, 0, 0, 0)
        cable_layout.setSpacing(4)

        self.cable_label = QLabel()
        self.cable_label.setStyleSheet("font-weight: bold;")
        cable_layout.addWidget(self.cable_label)

        self.cable_info_table = QTableWidget(0, 2)
        self.cable_info_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.cable_info_table.setSelectionMode(QAbstractItemView.NoSelection)
        self.cable_info_table.verticalHeader().setVisible(False)
        self.cable_info_table.horizontalHeader().setStretchLastSection(True)
        self.cable_info_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.cable_info_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        cable_layout.addWidget(self.cable_info_table, 1)

        self.cable_container.setMinimumWidth(320)
        self.splitter.addWidget(self.cable_container)

        # 创建图表显示区域（初始隐藏）
        self.chart_container = QWidget()
        self.chart_container.setVisible(False)
        chart_layout = QVBoxLayout(self.chart_container)
        chart_layout.setContentsMargins(0, 0, 0, 0)
        chart_layout.setSpacing(0)

        # 设置图表背景为黑色（先导入QBrush）
        from PySide6.QtGui import QColor, QBrush

        # 创建统一的图表，显示三条曲线
        unified_chart = QChart()
        self.unified_chart = unified_chart
        # 删除标题
        unified_chart.setTitle("")  # 设置空标题
        unified_chart.legend().setVisible(True)
        unified_chart.legend().setAlignment(Qt.AlignTop)
        # 完全隐藏标题区域
        unified_chart.layout().setContentsMargins(0, -20, 0, 0)  # 使用负边距移除标题区域
        # 或者尝试通过设置非常小的标题来避免占据空间
        unified_chart.setTitleFont(QFont())  # 使用默认字体
        unified_chart.setTitleBrush(QBrush(QColor(30, 30, 35)))  # 设置标题颜色与背景相同

        unified_chart.setBackgroundBrush(QBrush(QColor(30, 30, 35)))  # 与Base背景色一致
        unified_chart.setPlotAreaBackgroundBrush(QBrush(QColor(25, 25, 30)))  # 绘图区域稍微暗一些
        unified_chart.setPlotAreaBackgroundVisible(True)

        # 设置图例样式（无背景、无边框）
        legend = unified_chart.legend()
        legend.setLabelColor(QColor(255, 255, 255))  # 白色文字
        legend.setBackgroundVisible(False)  # 移除背景
        legend.setPen(QColor(0, 0, 0, 0))  # 设置透明边框

        # 创建三条曲线，设置不同颜色（使用更鲜艳的颜色在黑色背景上更显眼）
        self.voltage_series = QLineSeries()
        self.voltage_series.setName(self._text("chart_voltage"))
        self.voltage_series.setColor(QColor(120, 200, 255))  # 亮蓝色
        unified_chart.addSeries(self.voltage_series)

        self.current_series = QLineSeries()
        self.current_series.setName(self._text("chart_current"))
        self.current_series.setColor(QColor(255, 170, 80))  # 亮橙色
        unified_chart.addSeries(self.current_series)

        self.power_series = QLineSeries()
        self.power_series.setName(self._text("chart_power"))
        self.power_series.setColor(QColor(100, 255, 100))  # 亮绿色
        unified_chart.addSeries(self.power_series)

        # X轴 - 相对时间（秒）
        self.unified_axis_x = QValueAxis()
        self.unified_axis_x.setTitleText(self._text("chart_axis_x"))
        self.unified_axis_x.setTitleBrush(QBrush(QColor(255, 255, 255)))  # 白色标题
        self.unified_axis_x.setLabelsBrush(QBrush(QColor(200, 200, 200)))  # 浅灰色标签
        self.unified_axis_x.setGridLineColor(QColor(70, 70, 75))  # 深灰色网格线
        self.unified_axis_x.setLinePen(QColor(100, 100, 105))  # 坐标轴线颜色
        unified_chart.addAxis(self.unified_axis_x, Qt.AlignBottom)
        self.voltage_series.attachAxis(self.unified_axis_x)
        self.current_series.attachAxis(self.unified_axis_x)
        self.power_series.attachAxis(self.unified_axis_x)

        # Y轴 - 由于单位不同（V, A, W），使用一个统一的Y轴标签
        self.unified_axis_y = QValueAxis()
        self.unified_axis_y.setTitleText(self._text("chart_axis_y"))
        self.unified_axis_y.setTitleBrush(QBrush(QColor(255, 255, 255)))  # 白色标题
        self.unified_axis_y.setLabelsBrush(QBrush(QColor(200, 200, 200)))  # 浅灰色标签
        self.unified_axis_y.setGridLineColor(QColor(70, 70, 75))  # 深灰色网格线
        self.unified_axis_y.setLinePen(QColor(100, 100, 105))  # 坐标轴线颜色
        unified_chart.addAxis(self.unified_axis_y, Qt.AlignLeft)
        self.voltage_series.attachAxis(self.unified_axis_y)
        self.current_series.attachAxis(self.unified_axis_y)
        self.power_series.attachAxis(self.unified_axis_y)

        self.unified_chart_view = QChartView(unified_chart)
        self.unified_chart_view.setRenderHint(QPainter.Antialiasing)
        self.unified_chart_view.setBackgroundBrush(QBrush(QColor(30, 30, 35)))  # 视图背景
        chart_layout.addWidget(self.unified_chart_view, 1)

        self.splitter.addWidget(self.chart_container)

        self.splitter.setStretchFactor(0, 4)  # PD记录列表
        self.splitter.setStretchFactor(1, 3)  # 详细信息
        self.splitter.setStretchFactor(2, 2)  # PDO容器
        self.splitter.setStretchFactor(3, 3)  # 线缆信息
        self.splitter.setStretchFactor(4, 8)  # 图表区域（初始隐藏）
        self.splitter.setSizes([520, 360, 280, 340, 0])  # 初始时图表区域宽度为0

        self.summary_metrics = QFontMetrics(self.table.font())
        if self.summary_metrics:
            self.summary_char_px = max(6, self.summary_metrics.horizontalAdvance("0"))
            self.summary_line_height = max(16, self.summary_metrics.lineSpacing())
        base_height = self.summary_line_height + self.summary_base_padding
        self.table.verticalHeader().setDefaultSectionSize(base_height)

        status_bar = self.statusBar()

        self.status_label = QLabel()
        self.status_label.setStyleSheet(f"color: {STATUS_COLORS['status_disconnected']};")
        status_bar.addWidget(self.status_label)

        status_bar.addPermanentWidget(QLabel(" | "))

        self.count_label = QLabel()
        status_bar.addPermanentWidget(self.count_label)

        status_bar.addPermanentWidget(QLabel(" | "))

        self.measurement_count_label = QLabel()
        status_bar.addPermanentWidget(self.measurement_count_label)

        status_bar.addPermanentWidget(QLabel(" | "))

        self.voltage_label = QLabel("--")
        self.voltage_label.setToolTip(self._text("tooltip_voltage"))
        status_bar.addPermanentWidget(self.voltage_label)

        status_bar.addPermanentWidget(QLabel(" | "))

        self.current_label = QLabel("--")
        self.current_label.setToolTip(self._text("tooltip_current"))
        status_bar.addPermanentWidget(self.current_label)

        status_bar.addPermanentWidget(QLabel(" | "))

        self.power_label = QLabel("--")
        self.power_label.setToolTip(self._text("tooltip_power"))
        status_bar.addPermanentWidget(self.power_label)

        status_bar.addPermanentWidget(QLabel(" | "))

        self.auto_pause_status_label = QLabel(self._text("auto_pause_status_disabled"))
        self.auto_pause_status_label.setToolTip(self._text("auto_pause_status_tooltip"))
        status_bar.addPermanentWidget(self.auto_pause_status_label)

    def _update_measurement_display(self, voltage: Optional[float], current: Optional[float], power: Optional[float]) -> None:
        """Update voltage, current, and power displays in status bar."""
        if hasattr(self, "voltage_label"):
            if voltage is not None:
                text = self._text("measurement_display", value=voltage, unit="V")
                self.voltage_label.setText(text)
            else:
                self.voltage_label.setText("--")

        if hasattr(self, "current_label"):
            if current is not None:
                text = self._text("measurement_display", value=current, unit="A")
                self.current_label.setText(text)
            else:
                self.current_label.setText("--")

        if hasattr(self, "power_label"):
            if power is not None:
                text = self._text("measurement_display", value=power, unit="W")
                self.power_label.setText(text)
            else:
                self.power_label.setText("--")

    def _populate_device_list(self, _checked: bool = False) -> None:
        if not hasattr(self, "device_selector"):
            return
        if not HID_AVAILABLE:
            self.device_candidates = []
            self.device_selector.blockSignals(True)
            self.device_selector.clear()
            self.device_selector.addItem(self._text("device_selector_unavailable"))
            self.device_selector.blockSignals(False)
            self.device_selector.setEnabled(False)
            if hasattr(self, "refresh_devices_btn"):
                self.refresh_devices_btn.setEnabled(False)
            if hasattr(self, "connect_btn"):
                self.connect_btn.setEnabled(False)
            self.selected_device_value = None
            return

        previous_value = self.selected_device_value
        entries = self._get_available_devices()
        self.device_candidates = entries

        self.device_selector.blockSignals(True)
        self.device_selector.clear()

        selected_index = 0
        for idx, entry in enumerate(entries):
            label = entry.get("label") or self._text("device_fallback_label", index=idx + 1)
            self.device_selector.addItem(label)
            tooltip = entry.get("tooltip")
            if tooltip:
                self.device_selector.setItemData(idx, tooltip, Qt.ToolTipRole)
            value = entry.get("value")
            if previous_value is not None and value == previous_value:
                selected_index = idx

        if previous_value is None and len(entries) > 1:
            selected_index = 1

        self.device_selector.blockSignals(False)
        self.device_selector.setEnabled(not self.device_open)
        if hasattr(self, "refresh_devices_btn"):
            self.refresh_devices_btn.setEnabled(not self.device_open)
        if hasattr(self, "connect_btn"):
            self.connect_btn.setEnabled(True)

        self.device_selector.setCurrentIndex(selected_index)
        self._on_device_selected(selected_index)

    def _get_available_devices(self) -> List[Dict[str, Any]]:
        entries: List[Dict[str, Any]] = [
            {
                "label": self._text("device_auto"),
                "value": None,
                "tooltip": self._text("tooltip_default_vidpid", vid=K2_TARGET_VID, pid=K2_TARGET_PID),
            }
        ]

        if not HID_AVAILABLE:
            return entries

        try:
            enumerated = list(enumerate_devices())
        except Exception as exc:
            print(self._text("log_get_devices_failed", detail=exc))
            enumerated = []

        for dev in enumerated or []:
            path = dev.get("path")
            if not path:
                continue

            vid = dev.get("vendor_id", K2_TARGET_VID)
            pid = dev.get("product_id", K2_TARGET_PID)
            manufacturer = dev.get("manufacturer_string") or dev.get("manufacturer")
            product = dev.get("product_string") or dev.get("product")
            serial = dev.get("serial_number") or dev.get("serial")

            label_parts: List[str] = []
            if product:
                label_parts.append(str(product))
            elif manufacturer:
                label_parts.append(str(manufacturer))
            if serial:
                label_parts.append(str(serial))
            label = " - ".join(label_parts) if label_parts else self._text("device_fallback_label", index=len(entries))

            tooltip_lines = [self._text("tooltip_vid_pid", vid=vid, pid=pid)]
            if manufacturer:
                tooltip_lines.append(self._text("tooltip_manufacturer", manufacturer=manufacturer))
            if product:
                tooltip_lines.append(self._text("tooltip_product", product=product))
            if serial:
                tooltip_lines.append(self._text("tooltip_serial", serial=serial))
            if isinstance(path, (bytes, bytearray)):
                try:
                    path_display = path.decode("utf-8")
                except Exception:
                    path_display = str(path)
            else:
                path_display = str(path)
            tooltip_lines.append(self._text("tooltip_path", path=path_display))

            entries.append({
                "label": label,
                "value": {
                    "path": path,
                    "vid": vid,
                    "pid": pid,
                    "serial": serial,
                    "manufacturer": manufacturer,
                    "product": product,
                },
                "tooltip": "\n".join(str(line) for line in tooltip_lines if line),
            })

        return entries

    def _on_device_selected(self, index: int) -> None:
        if index < 0 or index >= len(self.device_candidates):
            self.selected_device_value = None
            return
        self.selected_device_value = self.device_candidates[index].get("value")

    def _get_selected_device_info(self) -> Any:
        if not hasattr(self, "device_selector"):
            return None
        index = self.device_selector.currentIndex()
        if 0 <= index < len(self.device_candidates):
            return self.device_candidates[index].get("value")
        return None

    def _on_language_checkbox_toggled(self, checked: bool) -> None:
        target_language = "en" if checked else "zh"
        if target_language == self.current_language:
            return
        self.current_language = target_language
        self._apply_language()
        if not self.device_open:
            self._populate_device_list()
        else:
            self._refresh_device_selector_labels()

    def _on_data_visualization_toggled(self, checked: bool) -> None:
        if not hasattr(self, "chart_container") or not hasattr(self, "splitter"):
            return

        # 显示/隐藏图表容器
        self.chart_container.setVisible(checked)

        # 显示/隐藏PDO和线缆信息容器
        if hasattr(self, "pdo_container"):
            self.pdo_container.setVisible(not checked)
        if hasattr(self, "cable_container"):
            self.cable_container.setVisible(not checked)

        # 更新splitter的大小
        if checked:
            # 显示图表，隐藏PDO和线缆信息
            self.splitter.setSizes([520, 360, 0, 0, 900])
            self._update_charts()
        else:
            # 隐藏图表，显示PDO和线缆信息
            self.splitter.setSizes([520, 360, 280, 340, 0])

    def _toggle_connection(self):
        if not self.device_open:
            self._connect_device()
        else:
            self._disconnect_device()

    def _connect_device(self):
        if not HID_AVAILABLE:
            QMessageBox.critical(self, self._text("dialog_error_title"), self._text("hid_unavailable"))
            return

        if HID_AVAILABLE and not self.device_candidates:
            self._populate_device_list()

        if self.log_records:
            reply = QMessageBox.question(
                self,
                self._text("dialog_clear_existing_title"),
                self._text("dialog_clear_existing"),
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                return
            self._clear_records(confirm=False)

        device_info = self._get_selected_device_info() if HID_AVAILABLE else None

        try:
            self.data_queue = Queue(maxsize=1000)
            self.stop_event = Event()
            self.pause_flag = Value('i', 1)

            self.collection_process = Process(
                target=data_collection_worker,
                args=(self.data_queue, self.stop_event, self.pause_flag, device_info),
                daemon=True
            )
            self.collection_process.start()

            self.device_open = True
            if hasattr(self, "start_btn"):
                self.start_btn.setEnabled(True)
            if HID_AVAILABLE and hasattr(self, "device_selector"):
                self.device_selector.setEnabled(False)
            if HID_AVAILABLE and hasattr(self, "refresh_devices_btn"):
                self.refresh_devices_btn.setEnabled(False)
            self._update_connect_button_text()
            self._update_start_button_text()
            self._set_status_message("status_connected")
        except Exception as exc:
            if HID_AVAILABLE and hasattr(self, "device_selector"):
                self.device_selector.setEnabled(True)
            if HID_AVAILABLE and hasattr(self, "refresh_devices_btn"):
                self.refresh_devices_btn.setEnabled(True)
            QMessageBox.critical(self, self._text("dialog_error_title"), self._text("error_connection_failed", detail=exc))

    def _disconnect_device(self):
        if self.stop_event:
            self.stop_event.set()
        if self.collection_process:
            self.collection_process.join(timeout=2)
            if self.collection_process.is_alive():
                self.collection_process.terminate()

        if self.data_queue:
            try:
                while True:
                    self.data_queue.get_nowait()
            except queue.Empty:
                pass

        self.device_open = False
        self.is_paused = True
        self._cancel_auto_pause_delay()
        self._update_connect_button_text()
        self._update_start_button_text()
        if hasattr(self, "start_btn"):
            self.start_btn.setEnabled(False)
        self._set_status_message("status_disconnected")

        # Reset measurement display
        self._update_measurement_display(None, None, None)

        if HID_AVAILABLE and hasattr(self, "refresh_devices_btn"):
            self.refresh_devices_btn.setEnabled(True)
        if HID_AVAILABLE and hasattr(self, "device_selector"):
            self.device_selector.setEnabled(True)
            self._populate_device_list()

    def _toggle_collection(self):
        if not self.device_open:
            return

        self._cancel_auto_pause_delay()

        if self.is_paused:
            self.is_paused = False
            if self.pause_flag:
                self.pause_flag.value = 0
            if self.start_time is None:
                self.start_time = time.time()
            self._set_status_message("status_collecting")
        else:
            self.is_paused = True
            if self.pause_flag:
                self.pause_flag.value = 1
            self._set_status_message("status_paused")
        self._update_start_button_text()

    def _clear_records(self, confirm: bool = True) -> bool:
        if confirm:
            reply = QMessageBox.question(
                self,
                self._text("dialog_confirm_title"),
                self._text("dialog_confirm_clear"),
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                return False

        self._reset_records_state()
        return True

    def _confirm_manual_clear(self) -> None:
        reply = QMessageBox.warning(
            self,
            self._text("dialog_manual_clear_title"),
            self._text("dialog_manual_clear"),
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        self._clear_records(confirm=False)

    def _reset_records_state(self) -> None:
        self.log_records.clear()
        self.log_index = 0
        self.measurement_records.clear()
        self.measurement_index = 0

        # 清空原始完整记录
        self.raw_log_records.clear()
        self.raw_measurement_records.clear()

        # 清空待处理的记录
        self.pending_records.clear()
        self.pending_measurements.clear()

        self.start_time = None
        self.current_max_lines = 1
        if hasattr(self, "table"):
            self.table.setRowCount(0)
            base_height = self.summary_line_height + self.summary_base_padding
            self.table.verticalHeader().setDefaultSectionSize(base_height)
        if hasattr(self, "detail_text"):
            self.detail_text.clear()
        if hasattr(self, "current_pdo_table"):
            self.current_pdo_table.setRowCount(0)
        self._update_cable_info([])
        self._update_count_label()
        self._update_measurement_count_label()
        # Reset measurement display
        self._update_measurement_display(None, None, None)

    def _export_csv(self):
        if not self.raw_log_records and not self.raw_measurement_records:
            QMessageBox.warning(self, self._text("dialog_warning_title"), self._text("dialog_export_no_records"))
            return

        filename, _ = QFileDialog.getSaveFileName(
            self,
            self._text("dialog_export_title"),
            f"PD_Records_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            self._text("file_filter_csv")
        )
        if not filename:
            return

        try:
            with open(filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow([
                    self._text("csv_header_index"),
                    self._text("csv_header_absolute_time"),
                    self._text("csv_header_relative_time"),
                    self._text("csv_header_type"),
                    self._text("csv_header_summary"),
                    self._text("csv_header_details"),
                    "Voltage(V)",  # 添加电压列
                    "Current(A)",  # 添加电流列
                    "Power(W)",    # 添加功率列
                ])

                # 导出PD记录（从原始完整记录导出）
                for record in self.raw_log_records:
                    relative_time = f"{record.get('relative_time', 0):.3f}"
                    detail = ""

                    if record['type'] == 'PDO' and isinstance(record.get('data'), list):
                        detail_parts = []
                        for entry in record['data']:
                            detail_parts.append(
                                f"PDO{entry.get('index', '?')}: "
                                f"{entry.get('summary', '')} [{entry.get('raw', '')}]"
                            )
                        detail = " | ".join(detail_parts)
                    elif record['type'] == 'RDO' and isinstance(record.get('data'), dict):
                        parts = []
                        if record['data'].get('raw'):
                            parts.append(f"Raw: {record['data']['raw']}")
                        if record['data'].get('details'):
                            parts.append(record['data']['details'])
                        detail = " | ".join(parts) if parts else ""

                    writer.writerow([
                        record['index'],
                        record['timestamp'],
                        relative_time,
                        record['type'],
                        record['summary'],
                        detail
                    ])

                # 导出测量数据（从原始完整记录导出）
                for record in self.raw_measurement_records:
                    relative_time = f"{record.get('relative_time', 0):.3f}"
                    writer.writerow([
                        record['index'],
                        record['timestamp'],
                        relative_time,
                        record['type'],
                        f"V:{record['voltage']:.3f}V I:{record['current']:.3f}A P:{record['power']:.3f}W",
                        "",
                        f"{record['voltage']:.3f}",
                        f"{record['current']:.3f}",
                        f"{record['power']:.3f}",
                    ])

            total_records = len(self.raw_log_records) + len(self.raw_measurement_records)
            QMessageBox.information(self, self._text("dialog_export_success_title"), self._text("dialog_export_success", count=total_records))
        except Exception as exc:
            QMessageBox.critical(self, self._text("dialog_error_title"), self._text("dialog_export_failed", error=exc))

    def _import_csv(self):
        filename, _ = QFileDialog.getOpenFileName(
            self,
            self._text("dialog_import_title"),
            "",
            self._text("dialog_import_filter")
        )
        if not filename:
            return

        try:
            self._reset_records_state()

            imported_count = 0

            with open(filename, 'r', encoding='utf-8-sig') as csvfile:
                reader = csv.reader(csvfile)
                rows = list(reader)

            header_variants = [
                (
                    LANG_STRINGS['zh']["csv_header_index"],
                    LANG_STRINGS['zh']["csv_header_absolute_time"],
                    LANG_STRINGS['zh']["csv_header_relative_time"],
                    LANG_STRINGS['zh']["csv_header_type"],
                    LANG_STRINGS['zh']["csv_header_summary"],
                    LANG_STRINGS['zh']["csv_header_details"],
                ),
                (
                    LANG_STRINGS['en']["csv_header_index"],
                    LANG_STRINGS['en']["csv_header_absolute_time"],
                    LANG_STRINGS['en']["csv_header_relative_time"],
                    LANG_STRINGS['en']["csv_header_type"],
                    LANG_STRINGS['en']["csv_header_summary"],
                    LANG_STRINGS['en']["csv_header_details"],
                ),
            ]

            if rows and tuple(rows[0]) in header_variants:
                rows = rows[1:]

            for row in rows:
                if not row or len(row) < 5:
                    continue

                try:
                    index = int(row[0])
                    timestamp = row[1]
                    relative_time = float(row[2])
                    record_type = row[3]
                    summary = row[4]
                    detail_str = row[5] if len(row) > 5 else ''

                    # 处理测量数据（如果有电压电流功率列）
                    if len(row) > 8 and row[6] and row[7] and row[8]:
                        try:
                            voltage = float(row[6])
                            current = abs(float(row[7]))
                            power = float(row[8])

                            # 添加测量数据记录（添加到原始完整记录）
                            measurement_record = {
                                "index": index,
                                "timestamp": timestamp,
                                "relative_time": relative_time,
                                "type": "MEASUREMENT",
                                "voltage": voltage,
                                "current": current,
                                "power": power
                            }

                            self.raw_measurement_records.append(measurement_record)
                            # 检查UI记录限制
                            if len(self.measurement_records) < self.ui_measurement_limit:
                                self.measurement_records.append(measurement_record)

                            self.measurement_index = max(self.measurement_index, index)
                            imported_count += 1
                            continue  # 跳过PD记录处理，这是测量数据
                        except (ValueError, IndexError):
                            pass  # 如果转换失败，继续作为PD记录处理

                    data = None
                    if record_type == 'PDO' and detail_str:
                        data = []
                        entries = detail_str.split(' | ')
                        for entry in entries:
                            if entry.startswith('PDO'):
                                parts = entry.split(': ', 1)
                                if len(parts) == 2:
                                    pdo_index = parts[0].replace('PDO', '')
                                    rest = parts[1]
                                    if '[' in rest and ']' in rest:
                                        summary_part = rest[:rest.rindex('[')].strip()
                                        raw_part = rest[rest.rindex('[')+1:rest.rindex(']')]
                                        data.append({
                                            'index': pdo_index,
                                            'summary': summary_part,
                                            'raw': raw_part
                                        })
                    elif record_type == 'RDO' and detail_str:
                        data = {}
                        parts = detail_str.split(' | ')
                        for part in parts:
                            if part.startswith('Raw: '):
                                data['raw'] = part.replace('Raw: ', '')
                            elif part:
                                data['details'] = part

                    record = {
                        "index": index,
                        "timestamp": timestamp,
                        "relative_time": relative_time,
                        "type": record_type,
                        "summary": summary,
                        "data": data
                    }

                    self.raw_log_records.append(record)
                    if len(self.log_records) < self.ui_log_limit:
                        self.log_records.append(record)
                        # 添加到表格
                        self._add_record_to_table(record)

                    self.log_index = max(self.log_index, index)
                    imported_count += 1

                except Exception as exc:
                    print(self._text("log_skip_invalid_row", detail=exc))
                    continue

            self._update_count_label()
            self._update_measurement_count_label()  # 更新测量数据计数

            pd_count = len(self.log_records)
            measurement_count = len(self.measurement_records)
            total_count = pd_count + measurement_count

            if measurement_count > 0:
                QMessageBox.information(self, self._text("dialog_import_success_title"),
                    f"成功导入 {total_count} 条记录 (PD: {pd_count}, 测量: {measurement_count})")
            else:
                QMessageBox.information(self, self._text("dialog_import_success_title"), self._text("dialog_import_success", count=imported_count))

        except Exception as exc:
            QMessageBox.critical(self, self._text("dialog_error_title"), self._text("dialog_import_failed", error=exc))

    def _poll_queue(self):
        if not self.data_queue or not self.device_open:
            return

        try:
            while True:
                payload = self.data_queue.get_nowait()
                self._handle_payload(payload)
        except queue.Empty:
            pass
        except Exception as exc:
            print(f"处理数据错误: {exc}")

    def _batch_update_ui(self):
        """批量更新UI（每200ms执行一次，减少UI刷新频率）"""
        # 批量添加PD记录
        if self.pending_records:
            # 取出所有待处理的记录
            records_to_add = self.pending_records[:]
            self.pending_records.clear()

            # 添加到原始完整记录
            self.raw_log_records.extend(records_to_add)

            # 检查是否需要滚动（如果达到UI限制）
            if len(self.log_records) >= self.ui_log_limit:
                # 滚动记录：移除旧的记录以保持UI限制
                remove_count = len(self.log_records) + len(records_to_add) - self.ui_log_limit
                if remove_count > 0:
                    self.log_records = self.log_records[remove_count:]
                    # 移除表格中的旧行
                    self.table.setRowCount(max(0, self.table.rowCount() - remove_count))

            # 添加到UI记录列表
            self.log_records.extend(records_to_add)

            # 批量添加到表格
            for record in records_to_add:
                self._add_record_to_table(record)

            self._update_count_label()

        # 批量添加测量记录
        if self.pending_measurements:
            measurements_to_add = self.pending_measurements[:]
            self.pending_measurements.clear()

            # 添加到原始完整记录
            self.raw_measurement_records.extend(measurements_to_add)

            # 检查是否需要滚动测量记录
            if len(self.measurement_records) >= self.ui_measurement_limit:
                remove_count = len(self.measurement_records) + len(measurements_to_add) - self.ui_measurement_limit
                if remove_count > 0:
                    self.measurement_records = self.measurement_records[remove_count:]

            # 添加到UI记录列表
            self.measurement_records.extend(measurements_to_add)
            self._update_measurement_count_label()

            # 更新图表（如果有大量测量数据，限制更新频率）
            if hasattr(self, "data_visualization_checkbox") and self.data_visualization_checkbox.isChecked():
                self._update_charts()

    def _handle_payload(self, payload: Dict[str, Any]):
        error_info = payload.get("error")
        if error_info:
            self._show_error(str(error_info))
            return

        # Update measurements display if available (even when paused)
        measurements = payload.get("measurements")
        if measurements:
            voltage = measurements.get("voltage")
            if voltage is not None:
                try:
                    voltage = float(voltage)
                except Exception:
                    pass
            current = measurements.get("current")
            # Ensure we display and record absolute current values; ignore if current is None
            if current is not None:
                try:
                    current = abs(float(current))
                except Exception:
                    # If parsing fails, leave as-is (should not happen normally)
                    pass
            power = measurements.get("power")
            self._last_voltage = voltage
            self._last_current = current
            self._update_measurement_display(voltage, current, power)

            # 检查自动暂停阈值（只有在收集状态下才检查）
            if not self.is_paused and self.device_open:
                self._check_and_apply_auto_pause(voltage, current)

            # 记录测量数据（只有在收集状态下才记录）- 添加到待处理列表
            if not self.is_paused and self.device_open:
                self.measurement_index += 1
                timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
                relative_time = 0.0
                if self.start_time is not None:
                    relative_time = time.time() - self.start_time

                measurement_record = {
                    "index": self.measurement_index,
                    "timestamp": timestamp,
                    "relative_time": relative_time,
                    "type": "MEASUREMENT",
                    "voltage": voltage,
                    "current": current,
                    "power": power
                }
                self.pending_measurements.append(measurement_record)

        if self.is_paused:
            return

        pkg = payload.get("data")
        if pkg is None:
            return

        cable_rows = CableDataParser.parse(payload)
        if cable_rows:
            self._update_cable_info(cable_rows)

        if is_pdo_packet(pkg):
            pdo_data = PDParser.parse_pdo(payload)
            if pdo_data is not None:
                self._update_current_pdos(pdo_data)
            if pdo_data:
                summary = " | ".join(entry["summary"] for entry in pdo_data if entry["summary"].strip())
                if summary.strip():
                    self._add_record_to_pending(payload, "PDO", summary, pdo_data)

        if is_rdo_packet(pkg):
            rdo_info = PDParser.parse_rdo(payload)
            summary = rdo_info.get("summary", "")
            if summary and summary != "Invalid RDO":
                self._add_record_to_pending(payload, "RDO", summary, rdo_info)

    def _add_record_to_pending(self, payload: Dict[str, Any], record_type: str, summary: str, data: Any):
        """将记录添加到待处理列表，等待批量更新"""
        self.log_index += 1

        timestamp = payload.get("timestamp", datetime.now().strftime("%H:%M:%S.%f")[:-3])
        time_sec = payload.get("time_sec", time.time())

        relative_time = 0.0
        if self.start_time is not None:
            relative_time = time_sec - self.start_time

        record = {
            "index": self.log_index,
            "timestamp": timestamp,
            "relative_time": relative_time,
            "type": record_type,
            "summary": summary,
            "data": data
        }

        self.pending_records.append(record)

    def _add_record_to_table(self, record: Dict[str, Any]):
        display_summary = self._wrap_summary(record["summary"])
        num_lines = display_summary.count('\n') + 1
        self.current_max_lines = max(self.current_max_lines, num_lines)

        new_height = self.summary_line_height * self.current_max_lines + self.summary_base_padding
        self.table.verticalHeader().setDefaultSectionSize(new_height)

        row = self.table.rowCount()
        self.table.insertRow(row)

        index_item = QTableWidgetItem(str(record["index"]))
        index_item.setTextAlignment(Qt.AlignCenter)
        time_item = QTableWidgetItem(f"{record['relative_time']:.3f}")
        time_item.setTextAlignment(Qt.AlignCenter)
        type_item = QTableWidgetItem(record["type"])
        type_item.setTextAlignment(Qt.AlignCenter)
        summary_item = QTableWidgetItem(display_summary)
        summary_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        for item in (index_item, time_item, type_item, summary_item):
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)

        self.table.setItem(row, 0, index_item)
        self.table.setItem(row, 1, time_item)
        self.table.setItem(row, 2, type_item)
        self.table.setItem(row, 3, summary_item)

        if hasattr(self, "auto_scroll_checkbox") and self.auto_scroll_checkbox.isChecked():
            self.table.scrollToBottom()
        self.table.resizeRowToContents(row)

    def _update_count(self):
        self._update_count_label()

    def _update_charts(self) -> None:
        """更新统一图表，显示测量数据"""
        if not hasattr(self, "voltage_series"):
            return

        # 清空现有数据
        self.voltage_series.clear()
        self.current_series.clear()
        self.power_series.clear()

        if not self.measurement_records:
            return

        # 限制显示的数据点数量，避免性能问题
        max_points = 1000
        records = self.measurement_records
        if len(records) > max_points:
            # 取最新的max_points个点
            step = len(records) // max_points + 1
            records = records[::step]

        # 添加数据点到图表
        min_voltage = float('inf')
        max_voltage = float('-inf')
        min_current = float('inf')
        max_current = float('-inf')
        min_power = float('inf')
        max_power = float('-inf')

        # 使用相对时间作为X轴数据
        for record in records:
            relative_seconds = record.get('relative_time', 0)
            voltage = record.get('voltage', 0)
            current = record.get('current', 0)
            power = record.get('power', 0)

            self.voltage_series.append(relative_seconds, voltage)
            self.current_series.append(relative_seconds, current)
            self.power_series.append(relative_seconds, power)

            # 计算所有值的最小最大值
            min_voltage = min(min_voltage, voltage)
            max_voltage = max(max_voltage, voltage)
            min_current = min(min_current, current)
            max_current = max(max_current, current)
            min_power = min(min_power, power)
            max_power = max(max_power, power)

        # 设置X轴范围（相对时间）
        if records:
            first_relative = records[0].get('relative_time', 0)
            last_relative = records[-1].get('relative_time', 0)

            self.unified_axis_x.setRange(first_relative, last_relative)

        # 设置统一的Y轴范围（覆盖所有值）
        if hasattr(self, 'unified_axis_y'):
            all_min = min(min_voltage, min_current, min_power)
            all_max = max(max_voltage, max_current, max_power)

            if all_min != float('inf') and all_max != float('-inf'):
                margin = (all_max - all_min) * 0.1
                self.unified_axis_y.setRange(all_min - margin, all_max + margin)

    def _apply_auto_pause_settings(
        self,
        enabled: bool,
        metric: str,
        voltage_threshold: float,
        current_threshold: float,
        delay_seconds: float,
    ) -> None:
        """Persist auto pause settings after validation."""
        metric_key = metric if metric in {"voltage", "current"} else "voltage"

        self.auto_pause_threshold_enabled = enabled
        self.auto_pause_metric = metric_key
        self.voltage_threshold = voltage_threshold
        self.current_threshold = current_threshold
        self.auto_pause_delay_seconds = max(0.0, delay_seconds)

        if not enabled:
            self._cancel_auto_pause_delay()

        self._update_auto_pause_status_display()

    def _show_auto_pause_settings(self) -> None:
        """显示自动暂停设置对话框"""
        from PySide6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLineEdit, QDialogButtonBox

        dialog = QDialog(self)
        dialog.setWindowTitle(self._text("dialog_threshold_title"))
        layout = QVBoxLayout(dialog)

        hint_label = QLabel(self._text("threshold_settings_hint"))
        hint_label.setWordWrap(True)
        layout.addWidget(hint_label)

        enable_checkbox = QCheckBox(self._text("checkbox_auto_pause_threshold"))
        enable_checkbox.setChecked(self.auto_pause_threshold_enabled)
        layout.addWidget(enable_checkbox)

        metric_layout = QHBoxLayout()
        metric_label = QLabel(self._text("label_auto_pause_metric"))
        metric_combo = QComboBox()
        metric_combo.addItem(self._text("auto_pause_metric_voltage"), "voltage")
        metric_combo.addItem(self._text("auto_pause_metric_current"), "current")
        metric_combo.setCurrentIndex(0 if self.auto_pause_metric != "current" else 1)
        metric_layout.addWidget(metric_label)
        metric_layout.addWidget(metric_combo)
        layout.addLayout(metric_layout)

        voltage_layout = QHBoxLayout()
        voltage_label = QLabel(self._text("label_voltage_threshold"))
        voltage_input = QLineEdit(str(self.voltage_threshold))
        voltage_input.setToolTip(self._text("tooltip_voltage_threshold"))
        voltage_layout.addWidget(voltage_label)
        voltage_layout.addWidget(voltage_input)
        layout.addLayout(voltage_layout)

        current_layout = QHBoxLayout()
        current_label = QLabel(self._text("label_current_threshold"))
        current_input = QLineEdit(str(self.current_threshold))
        current_input.setToolTip(self._text("tooltip_current_threshold"))
        current_layout.addWidget(current_label)
        current_layout.addWidget(current_input)
        layout.addLayout(current_layout)

        delay_layout = QHBoxLayout()
        delay_label = QLabel(self._text("label_auto_pause_delay"))
        delay_input = QLineEdit(str(self.auto_pause_delay_seconds))
        delay_input.setToolTip(self._text("tooltip_auto_pause_delay"))
        delay_layout.addWidget(delay_label)
        delay_layout.addWidget(delay_input)
        layout.addLayout(delay_layout)

        def on_enable_toggled(checked: bool) -> None:
            metric_combo.setEnabled(checked)
            voltage_input.setEnabled(checked)
            current_input.setEnabled(checked)
            delay_input.setEnabled(checked)

        on_enable_toggled(self.auto_pause_threshold_enabled)
        enable_checkbox.toggled.connect(on_enable_toggled)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)

        if dialog.exec() != QDialog.Accepted:
            return

        try:
            voltage_threshold = float(voltage_input.text())
            current_threshold = float(current_input.text())
            delay_seconds = float(delay_input.text())
        except ValueError:
            QMessageBox.warning(self, self._text("dialog_warning_title"), self._text("auto_pause_invalid_number"))
            return

        if voltage_threshold < 0 or voltage_threshold > 20:
            QMessageBox.warning(self, self._text("dialog_warning_title"), self._text("auto_pause_voltage_range_warning"))
            return

        if current_threshold < 0 or current_threshold > 5:
            QMessageBox.warning(self, self._text("dialog_warning_title"), self._text("auto_pause_current_range_warning"))
            return

        if delay_seconds < 0 or delay_seconds > 10:
            QMessageBox.warning(self, self._text("dialog_warning_title"), self._text("auto_pause_delay_range_warning"))
            return

        enabled = enable_checkbox.isChecked()
        metric = metric_combo.currentData() or "voltage"

        self._apply_auto_pause_settings(
            enabled=enabled,
            metric=metric,
            voltage_threshold=voltage_threshold,
            current_threshold=current_threshold,
            delay_seconds=delay_seconds,
        )

        if enabled:
            metric_label = self._text("auto_pause_metric_voltage" if metric == "voltage" else "auto_pause_metric_current")
            unit = "V" if metric == "voltage" else "A"
            threshold_value = voltage_threshold if metric == "voltage" else current_threshold
            info_message = self._text(
                "auto_pause_settings_enabled_summary",
                metric=metric_label,
                threshold=f"{threshold_value:.3f}",
                unit=unit,
                delay=f"{delay_seconds:.1f}",
            )
        else:
            info_message = self._text("auto_pause_settings_disabled")

        QMessageBox.information(self, self._text("dialog_info_title"), info_message)

    def _update_auto_pause_status_display(self) -> None:
        """更新状态栏的自动暂停状态显示"""
        if not hasattr(self, "auto_pause_status_label"):
            return

        if self.auto_pause_threshold_enabled:
            metric = self.auto_pause_metric if self.auto_pause_metric in {"voltage", "current"} else "voltage"
            threshold = self.voltage_threshold if metric == "voltage" else self.current_threshold
            metric_label = self._text("auto_pause_metric_voltage" if metric == "voltage" else "auto_pause_metric_current")
            unit = "V" if metric == "voltage" else "A"
            delay_text = ""
            if self.auto_pause_delay_seconds > 0:
                delay_text = self._text(
                    "auto_pause_status_delay_suffix",
                    delay=f"{self.auto_pause_delay_seconds:.1f}"
                )
            status_text = self._text(
                "auto_pause_status_enabled_single",
                metric=metric_label,
                threshold=threshold,
                unit=unit,
                delay_text=delay_text,
            )
            self.auto_pause_status_label.setText(status_text)
            self.auto_pause_status_label.setStyleSheet("color: #4fd3c4;")
        else:
            self.auto_pause_status_label.setText(self._text("auto_pause_status_disabled"))
            self.auto_pause_status_label.setStyleSheet("color: #bbbbbb;")

    def _check_and_apply_auto_pause(self, voltage: Optional[float], current: Optional[float]) -> None:
        """检查是否需要自动暂停"""
        if not self.auto_pause_threshold_enabled or not self.device_open or self.is_paused:
            self._cancel_auto_pause_delay()
            return

        metric = self.auto_pause_metric if self.auto_pause_metric in {"voltage", "current"} else "voltage"
        threshold = self.voltage_threshold if metric == "voltage" else self.current_threshold
        value = voltage if metric == "voltage" else current

        if value is None:
            self._cancel_auto_pause_delay()
            return

        if value < threshold:
            if self.auto_pause_delay_seconds <= 0:
                self._apply_auto_pause_trigger(metric, value, threshold)
            else:
                self._schedule_auto_pause_delay(metric, threshold)
        else:
            self._cancel_auto_pause_delay()

    def _schedule_auto_pause_delay(self, metric: str, threshold: float) -> None:
        """启动延时自动暂停计时器"""
        delay_ms = int(self.auto_pause_delay_seconds * 1000)
        if delay_ms <= 0:
            return

        if self.auto_pause_delay_timer.isActive():
            return

        self._auto_pause_pending_metric = metric
        self._auto_pause_pending_threshold = threshold
        self.auto_pause_delay_timer.start(delay_ms)

    def _execute_pending_auto_pause(self) -> None:
        """在延时结束后再次确认是否需要自动暂停"""
        metric = self._auto_pause_pending_metric
        threshold = self._auto_pause_pending_threshold
        self._auto_pause_pending_metric = None
        self._auto_pause_pending_threshold = None

        if not metric or threshold is None:
            return

        current_value = self._last_voltage if metric == "voltage" else self._last_current
        if current_value is None:
            return

        if current_value < threshold:
            self._apply_auto_pause_trigger(metric, current_value, threshold)

    def _cancel_auto_pause_delay(self) -> None:
        if self.auto_pause_delay_timer.isActive():
            self.auto_pause_delay_timer.stop()
        self._auto_pause_pending_metric = None
        self._auto_pause_pending_threshold = None

    def _apply_auto_pause_trigger(self, metric: str, value: Optional[float], threshold: float) -> None:
        """暂停采集并给出提示"""
        if self.is_paused:
            return

        self._cancel_auto_pause_delay()
        self._toggle_collection()

        self._set_status_message("status_paused", update_style=False)
        if hasattr(self, "status_label"):
            self.status_label.setStyleSheet("color: #ff5555;")

        metric_label = self._text("auto_pause_metric_voltage" if metric == "voltage" else "auto_pause_metric_current")
        unit = "V" if metric == "voltage" else "A"
        threshold_text = f"{threshold:.3f}"

        if value is None:
            message = self._text(
                "auto_pause_triggered_without_value",
                metric=metric_label,
                threshold=threshold_text,
                unit=unit,
            )
        else:
            value_text = f"{value:.3f}"
            message = self._text(
                "auto_pause_triggered_with_value",
                metric=metric_label,
                value=value_text,
                threshold=threshold_text,
                unit=unit,
            )

        QMessageBox.information(self, self._text("dialog_info_title"), message)

    def _update_current_pdos(self, entries: List[Dict[str, str]]) -> None:
        if not hasattr(self, "current_pdo_table"):
            return

        self.current_pdo_table.setRowCount(len(entries))
        for row, entry in enumerate(entries):
            index_item = QTableWidgetItem(entry.get("index", ""))
            index_item.setTextAlignment(Qt.AlignCenter)

            summary_text = entry.get("summary", "")
            summary_item = QTableWidgetItem(summary_text)

            for item in (index_item, summary_item):
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)

            self.current_pdo_table.setItem(row, 0, index_item)
            self.current_pdo_table.setItem(row, 1, summary_item)

        if entries:
            self.current_pdo_table.resizeRowsToContents()

    def _update_cable_info(self, entries: List[Tuple[str, Any]]) -> None:
        self.cable_info_rows = list(entries)
        if not hasattr(self, "cable_info_table"):
            return

        self.cable_info_table.setRowCount(len(entries))
        for row, (field_key, raw_value) in enumerate(entries):
            field_text = self._translate_cable_field(field_key)
            value_text = self._translate_cable_value(raw_value)

            field_item = QTableWidgetItem(field_text)
            field_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            field_item.setFlags(field_item.flags() & ~Qt.ItemIsEditable)

            value_item = QTableWidgetItem(value_text)
            value_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            value_item.setFlags(value_item.flags() & ~Qt.ItemIsEditable)

            self.cable_info_table.setItem(row, 0, field_item)
            self.cable_info_table.setItem(row, 1, value_item)

        if entries:
            self.cable_info_table.resizeRowsToContents()

    def _wrap_summary(self, text: str) -> str:
        if not text:
            return ""

        try:
            summary_column_width = max(120, self.table.columnWidth(3))
        except Exception:
            summary_column_width = 300

        char_px = max(6, self.summary_char_px)
        self.summary_wrap_chars = max(12, summary_column_width // char_px)

        effective_width = self.summary_wrap_chars
        max_lines = max(1, self.summary_max_lines)

        lines: List[str] = []
        current_line = ""
        truncated = False

        raw_segments: List[str] = []
        for raw_line in text.splitlines() or [text]:
            parts = [segment.strip() for segment in raw_line.split(" | ") if segment.strip()]
            if parts:
                raw_segments.extend(parts)

        if not raw_segments:
            return ""

        def flush_line():
            nonlocal current_line, truncated
            if current_line:
                lines.append(current_line)
            current_line = ""
            if len(lines) >= max_lines:
                truncated = True

        for segment in raw_segments:
            if truncated:
                break
            wrapped_parts = textwrap.wrap(
                segment,
                width=effective_width,
                break_long_words=False,
                replace_whitespace=False,
            ) or [segment]

            for part in wrapped_parts:
                if truncated:
                    break
                if not current_line:
                    current_line = part
                else:
                    candidate = f"{current_line} | {part}"
                    if len(candidate) <= effective_width:
                        current_line = candidate
                    else:
                        flush_line()
                        if truncated:
                            break
                        current_line = part

            if truncated:
                break

        if not truncated and current_line:
            lines.append(current_line)
        if len(lines) > max_lines:
            lines = lines[:max_lines]
            truncated = True

        if truncated and lines:
            lines[-1] = lines[-1].rstrip() + " ..."

        return "\n".join(lines)

    def _on_summary_column_resized(self, _index: int, _old: int, _new: int):
        if _index != 3 or self._updating_summary:
            return
        self._refresh_summary_wrapping()

    def _refresh_summary_wrapping(self):
        if self._updating_summary:
            return
        self._updating_summary = True
        try:
            items = self.table.rowCount()
            if items == 0:
                return

            max_lines_found = 1
            for row in range(items):
                index_item = self.table.item(row, 0)
                if not index_item:
                    continue
                try:
                    record_index = int(index_item.text())
                except ValueError:
                    continue
                record = next((r for r in self.log_records if r["index"] == record_index), None)
                if not record:
                    continue

                display_summary = self._wrap_summary(record["summary"])
                num_lines = display_summary.count('\n') + 1
                max_lines_found = max(max_lines_found, num_lines)

                summary_item = self.table.item(row, 3)
                if summary_item and summary_item.text() != display_summary:
                    summary_item.setText(display_summary)
                    self.table.resizeRowToContents(row)

            if max_lines_found != self.current_max_lines:
                self.current_max_lines = max_lines_found
                new_height = self.summary_line_height * max_lines_found + self.summary_base_padding
                self.table.verticalHeader().setDefaultSectionSize(new_height)
        finally:
            self._updating_summary = False

    def _on_selection_changed(self):
        selected = self.table.selectionModel().selectedRows()
        if not selected:
            return
        row = selected[0].row()
        index_item = self.table.item(row, 0)
        if not index_item:
            return
        try:
            record_index = int(index_item.text())
        except ValueError:
            return

        record = next((r for r in self.log_records if r['index'] == record_index), None)
        if not record:
            return

        details: List[str] = []
        details.append(f"{self._text('detail_index')}: {record['index']}")
        details.append(f"{self._text('detail_time')}: {record['timestamp']}")
        details.append(f"{self._text('detail_relative_time')}: {self._text('detail_relative_seconds', seconds=record['relative_time'])}")
        details.append(f"{self._text('detail_type')}: {record['type']}")
        details.append(f"{self._text('detail_summary')}: {record['summary']}")
        details.append("")

        if record['type'] == 'PDO' and isinstance(record.get('data'), list):
            details.append(self._text('detail_pdo_header'))
            for entry in record['data']:
                details.append(f"  PDO{entry.get('index', '?')}: {entry.get('summary', '')}")
                details.append(f"    {self._text('detail_raw')}: {entry.get('raw', '')}")
        elif record['type'] == 'RDO' and isinstance(record.get('data'), dict):
            details.append(self._text('detail_rdo_header'))
            if record['data'].get('raw'):
                details.append(f"  {self._text('detail_raw')}: {record['data']['raw']}")
            if record['data'].get('details'):
                details.append(f"  {self._text('detail_extra')}: {record['data']['details']}")

        self.detail_text.setPlainText("\n".join(details))

    def _show_error(self, message: str):
        display = message
        if message == "device_disconnected":
            display = self._text("error_device_disconnected")
        elif message.startswith("connection_failed"):
            parts = message.split(":", 1)
            if len(parts) == 2 and parts[1].strip():
                detail = parts[1].strip()
                display = self._text("error_connection_failed", detail=detail)
            else:
                display = self._text("error_connection_failed_generic")
        QMessageBox.critical(self, self._text("dialog_error_title"), display)
        if self.device_open:
            self._disconnect_device()

    def closeEvent(self, event):  # noqa: N802
        self.poll_timer.stop()
        self.ui_update_timer.stop()
        self.memory_check_timer.stop()
        if self.device_open:
            self._disconnect_device()
        super().closeEvent(event)


# ======================== Dark Theme ========================
def apply_dark_theme(app: QApplication):
    app.setStyle("Fusion")
    palette = QPalette()

    palette.setColor(QPalette.Window, QColor(40, 40, 45))
    palette.setColor(QPalette.WindowText, Qt.white)
    palette.setColor(QPalette.Base, QColor(30, 30, 35))
    palette.setColor(QPalette.AlternateBase, QColor(45, 45, 50))
    palette.setColor(QPalette.ToolTipBase, Qt.white)
    palette.setColor(QPalette.ToolTipText, Qt.white)
    palette.setColor(QPalette.Text, Qt.white)
    palette.setColor(QPalette.Button, QColor(45, 45, 50))
    palette.setColor(QPalette.ButtonText, Qt.white)
    palette.setColor(QPalette.Link, QColor(100, 180, 255))
    palette.setColor(QPalette.Highlight, QColor(90, 90, 95))
    palette.setColor(QPalette.HighlightedText, Qt.white)

    app.setPalette(palette)
    app.setStyleSheet("QToolTip { color: #ffffff; background-color: #2a82da; border: 1px solid white; }")


# ======================== 主程序入口 ========================
def main():
    freeze_support()
    app = QApplication(sys.argv)
    apply_dark_theme(app)

    window = PDViewerWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == '__main__':
    main()
