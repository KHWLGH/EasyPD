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
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFontMetrics, QPalette, QColor

from vendor_ids_dict import VENDOR_IDS

SUPPORTED_LANGUAGES = [
    ("zh", "中文"),
    ("en", "English"),
]

LANG_STRINGS = {
    "zh": {
        "app_title": "EasyPD",
        "label_device": "设备:",
        "label_language": "语言:",
        "btn_refresh_devices": "刷新设备",
        "btn_connect": "连接设备",
        "btn_disconnect": "断开设备",
        "btn_start": "开始收集",
        "btn_pause": "暂停收集",
        "btn_resume": "继续收集",
        "btn_clear": "清空记录",
        "btn_export": "导出 CSV",
        "btn_import": "导入 CSV",
        "checkbox_auto_scroll": "自动滚动",
        "status_disconnected": "未连接",
        "status_connected": "已连接",
        "status_collecting": "收集中...",
        "status_paused": "已暂停",
        "records_count": "记录: {count}",
        "status_records_zero": "记录: 0",
        "device_auto": "自动选择",
        "device_fallback_label": "设备 {index}",
        "device_selector_unavailable": "HID 库不可用",
        "hid_unavailable": "HID 库不可用，无法连接设备",
        "tooltip_manufacturer": "厂家: {manufacturer}",
        "tooltip_product": "产品: {product}",
        "tooltip_serial": "序列号: {serial}",
        "tooltip_path": "路径: {path}",
    "language_switch_tooltip": "切换至 {language}",
    "tooltip_default_vidpid": "默认 VID:PID = {vid:04X}:{pid:04X}",
    "tooltip_vid_pid": "VID:PID = {vid:04X}:{pid:04X}",
        "dialog_clear_existing": "检测到已有记录，重新连接前需要清空数据。是否清空？",
        "dialog_clear_existing_title": "清空记录",
        "dialog_confirm_clear": "确定要清空所有记录吗?",
        "dialog_confirm_title": "确认",
        "dialog_manual_clear": "此操作将立即清空所有记录且不可恢复，是否继续?",
        "dialog_manual_clear_title": "清空记录",
        "dialog_warning_title": "警告",
        "dialog_error_title": "错误",
        "dialog_info_title": "提示",
        "dialog_question_title": "询问",
        "dialog_export_title": "导出 CSV",
        "dialog_import_title": "导入 CSV",
        "dialog_export_no_records": "没有可导出的记录",
        "dialog_export_success": "成功导出 {count} 条记录",
        "dialog_export_success_title": "成功",
        "dialog_export_failed": "导出失败: {error}",
        "dialog_import_success": "成功导入 {count} 条记录",
        "dialog_import_success_title": "成功",
        "dialog_import_failed": "导入失败: {error}",
        "file_filter_csv": "CSV 文件 (*.csv);;所有文件 (*.*)",
        "csv_header_index": "序号",
        "csv_header_absolute_time": "绝对时间",
        "csv_header_relative_time": "相对时间(秒)",
        "csv_header_type": "类型",
        "csv_header_summary": "摘要",
        "csv_header_details": "详细数据",
        "pdo_title": "当前 PDO 列表",
        "pdo_column_index": "序号",
        "pdo_column_description": "描述",
        "detail_title": "详细信息",
        "cable_title": "线缆信息 (VDM)",
        "cable_column_field": "字段",
        "cable_column_value": "值",
        "table_column_index": "序号",
        "table_column_relative_time": "相对时间(秒)",
        "table_column_type": "类型",
        "table_column_summary": "摘要",
        "detail_index": "序号",
        "detail_time": "时间",
        "detail_relative_time": "相对时间",
        "detail_relative_seconds": "{seconds:.3f} 秒",
        "detail_type": "类型",
        "detail_summary": "摘要",
        "detail_pdo_header": "PDO 详细信息:",
        "detail_rdo_header": "RDO 详细信息:",
        "detail_raw": "原始数据",
        "detail_extra": "详情",
        "error_device_disconnected": "设备已断开连接",
        "error_connection_failed": "连接失败: {detail}",
        "error_connection_failed_generic": "连接失败",
        "language_option_zh": "中文",
        "language_option_en": "English",
        "detail_separator_seconds": "秒",
        "dialog_import_filter": "CSV 文件 (*.csv);;所有文件 (*.*)",
        "status_idle": "未连接",
        "log_get_devices_failed": "获取设备列表失败: {detail}",
        "log_skip_invalid_row": "跳过无效行: {detail}",
    },
    "en": {
        "app_title": "EasyPD",
        "label_device": "Device:",
        "label_language": "Language:",
        "btn_refresh_devices": "Refresh Devices",
        "btn_connect": "Connect",
        "btn_disconnect": "Disconnect",
        "btn_start": "Start Capture",
        "btn_pause": "Pause Capture",
        "btn_resume": "Resume Capture",
        "btn_clear": "Clear Records",
        "btn_export": "Export CSV",
        "btn_import": "Import CSV",
        "checkbox_auto_scroll": "Auto Scroll",
        "status_disconnected": "Disconnected",
        "status_connected": "Connected",
        "status_collecting": "Capturing...",
        "status_paused": "Paused",
        "records_count": "Records: {count}",
        "status_records_zero": "Records: 0",
        "device_auto": "Auto Select",
        "device_fallback_label": "Device {index}",
        "device_selector_unavailable": "HID library unavailable",
        "hid_unavailable": "HID library unavailable. Unable to connect.",
        "tooltip_manufacturer": "Manufacturer: {manufacturer}",
        "tooltip_product": "Product: {product}",
        "tooltip_serial": "Serial: {serial}",
        "tooltip_path": "Path: {path}",
    "language_switch_tooltip": "Switch to {language}",
    "tooltip_default_vidpid": "Default VID:PID = {vid:04X}:{pid:04X}",
    "tooltip_vid_pid": "VID:PID = {vid:04X}:{pid:04X}",
        "dialog_clear_existing": "Existing records detected. Clear data before reconnecting?",
        "dialog_clear_existing_title": "Clear Records",
        "dialog_confirm_clear": "Clear all records?",
        "dialog_confirm_title": "Confirm",
        "dialog_manual_clear": "This will erase all records immediately and cannot be undone. Continue?",
        "dialog_manual_clear_title": "Clear Records",
        "dialog_warning_title": "Warning",
        "dialog_error_title": "Error",
        "dialog_info_title": "Info",
        "dialog_question_title": "Question",
        "dialog_export_title": "Export CSV",
        "dialog_import_title": "Import CSV",
        "dialog_export_no_records": "No records available for export.",
        "dialog_export_success": "Exported {count} records successfully.",
        "dialog_export_success_title": "Success",
        "dialog_export_failed": "Export failed: {error}",
        "dialog_import_success": "Imported {count} records successfully.",
        "dialog_import_success_title": "Success",
        "dialog_import_failed": "Import failed: {error}",
        "file_filter_csv": "CSV Files (*.csv);;All Files (*.*)",
        "csv_header_index": "Index",
        "csv_header_absolute_time": "Absolute Time",
        "csv_header_relative_time": "Relative Time (s)",
        "csv_header_type": "Type",
        "csv_header_summary": "Summary",
        "csv_header_details": "Details",
        "pdo_title": "Current PDOs",
        "pdo_column_index": "Index",
        "pdo_column_description": "Description",
        "detail_title": "Details",
        "cable_title": "Cable Info (VDM)",
        "cable_column_field": "Field",
        "cable_column_value": "Value",
        "table_column_index": "Index",
        "table_column_relative_time": "Relative Time (s)",
        "table_column_type": "Type",
        "table_column_summary": "Summary",
        "detail_index": "Index",
        "detail_time": "Time",
        "detail_relative_time": "Relative Time",
        "detail_relative_seconds": "{seconds:.3f} s",
        "detail_type": "Type",
        "detail_summary": "Summary",
        "detail_pdo_header": "PDO Details:",
        "detail_rdo_header": "RDO Details:",
        "detail_raw": "Raw",
        "detail_extra": "Notes",
        "error_device_disconnected": "Device disconnected.",
        "error_connection_failed": "Connection failed: {detail}",
        "error_connection_failed_generic": "Connection failed.",
        "language_option_zh": "中文",
        "language_option_en": "English",
        "detail_separator_seconds": "s",
        "dialog_import_filter": "CSV Files (*.csv);;All Files (*.*)",
        "status_idle": "Disconnected",
        "log_get_devices_failed": "Failed to enumerate devices: {detail}",
        "log_skip_invalid_row": "Skipped invalid row: {detail}",
    },
}

CABLE_FIELD_TEXT = {
    "cable_svid": {"zh": "SVID", "en": "SVID"},
    "cable_source": {"zh": "来源", "en": "Source"},
    "cable_command": {"zh": "命令", "en": "Command"},
    "cable_vendor_id": {"zh": "厂商 VID", "en": "Vendor VID"},
    "cable_role": {"zh": "线缆角色", "en": "Cable Role"},
    "cable_product_id": {"zh": "产品 ID", "en": "Product ID"},
    "cable_device_version": {"zh": "设备版本", "en": "Device Version"},
    "cable_type": {"zh": "线缆类型", "en": "Cable Type"},
    "cable_connector": {"zh": "连接器", "en": "Connector"},
    "cable_termination": {"zh": "终端类型", "en": "Termination"},
    "cable_max_voltage": {"zh": "最大电压", "en": "Max Voltage"},
    "cable_current": {"zh": "电流能力", "en": "Current Capability"},
    "cable_highest_speed": {"zh": "最高速率", "en": "Highest Speed"},
    "cable_latency": {"zh": "线缆延迟", "en": "Latency"},
    "cable_supports_epr": {"zh": "支持 EPR", "en": "EPR Support"},
    "cable_sbu": {"zh": "SBU 线路", "en": "SBU Lines"},
    "cable_max_temp": {"zh": "最高工作温度", "en": "Max Operating Temp"},
    "cable_shutdown_temp": {"zh": "关断温度", "en": "Shutdown Temp"},
    "cable_usb4": {"zh": "USB4 支持", "en": "USB4 Support"},
    "cable_charge_through": {"zh": "支持穿透供电", "en": "Charge-Through Support"},
    "cable_charge_through_current": {"zh": "穿透电流能力", "en": "Charge-Through Current"},
    "cable_vbus_impedance": {"zh": "VBUS 阻抗", "en": "VBUS Impedance"},
    "cable_ground_impedance": {"zh": "地线阻抗", "en": "Ground Impedance"},
}

CABLE_VALUE_TEXT = {
    "cable_type_passive": {"zh": "被动线缆", "en": "Passive Cable"},
    "cable_type_active": {"zh": "主动线缆", "en": "Active Cable"},
    "cable_type_vpd": {"zh": "VCONN 供电设备", "en": "VCONN Powered Device"},
    "bool_true": {"zh": "是", "en": "Yes"},
    "bool_false": {"zh": "否", "en": "No"},
}

STATUS_COLORS = {
    "status_disconnected": "#bbbbbb",
    "status_connected": "#3fc1c9",
    "status_collecting": "#4fd3c4",
    "status_paused": "#ffa931",
}

# ======================== 设备通信模块 ========================
try:
    from witrnhid import WITRN_DEV, is_pdo, is_rdo
    from witrnhid.core import K2_TARGET_VID, K2_TARGET_PID
    try:
        import hid  # type: ignore
    except Exception:
        hid = None  # type: ignore
    HID_AVAILABLE = True
except ImportError:
    HID_AVAILABLE = False
    hid = None  # type: ignore
    print("警告：无法导入 witrnhid 模块，将使用模拟模式")

    class WITRN_DEV:  # type: ignore
        def __init__(self):
            pass

        def open(self, *args, **kwargs):
            raise Exception("HID 库不可用")

        def close(self):
            pass

        def read_data(self):
            pass

        def auto_unpack(self):
            return None, None

    def is_pdo(_pkg):
        return False

    def is_rdo(_pkg):
        return False
    K2_TARGET_VID = 0x0000
    K2_TARGET_PID = 0x0000


def is_sink_cap(pkg) -> bool:
    """检查是否为 Sink Capabilities 消息"""
    try:
        msg_type_value = pkg["Message Header"]["Message Type"].value()
        normalized = str(msg_type_value).strip().lower()
        return "sink" in normalized and "cap" in normalized
    except Exception:
        return False


def data_collection_worker(data_queue, stop_event, pause_flag, device_info=None):
    """数据采集工作进程"""
    k2 = WITRN_DEV()
    last_pdo = None
    last_rdo = None

    try:
        def attempt_open(*args, **kwargs) -> bool:
            try:
                k2.open(*args, **kwargs)
                return True
            except Exception:
                return False

        def try_open_with_info(info) -> bool:
            if info is None:
                return attempt_open()

            if isinstance(info, dict):
                path = info.get("path") or info.get("device_path")
                if isinstance(path, str):
                    try:
                        path_bytes = path.encode("utf-8", "ignore")
                    except Exception:
                        path_bytes = None
                else:
                    path_bytes = path

                if path_bytes and attempt_open(path_bytes):
                    return True

                vid = info.get("vid") if info.get("vid") is not None else info.get("vendor_id")
                pid = info.get("pid") if info.get("pid") is not None else info.get("product_id")
                try:
                    if vid is not None and pid is not None and attempt_open(int(vid), int(pid)):
                        return True
                except Exception:
                    pass

                if path and attempt_open(path=path):
                    return True

                return attempt_open()

            if isinstance(info, (bytes, bytearray)) and attempt_open(info):
                return True

            if isinstance(info, (list, tuple)) and len(info) == 2:
                try:
                    vid_val = int(info[0])
                    pid_val = int(info[1])
                except Exception:
                    pass
                else:
                    if attempt_open(vid_val, pid_val):
                        return True

            return attempt_open()

        if not try_open_with_info(device_info):
            raise Exception("无法打开所选设备")

        while not stop_event.is_set():
            try:
                k2.read_data()
                unpacked = k2.auto_unpack()

                if not unpacked or len(unpacked) != 2:
                    continue

                timestamp_str, pkg = unpacked
                if pkg is None:
                    continue

                field_type = getattr(pkg, "field", lambda: None)()
                if field_type != "pd":
                    continue

                if pause_flag.value == 1:
                    continue

                current_time = time.time()
                is_sink_cap_flag = is_sink_cap(pkg)
                is_pdo_flag = is_pdo(pkg) and not is_sink_cap_flag
                is_rdo_flag = is_rdo(pkg)

                if is_pdo_flag:
                    last_pdo = pkg
                if is_rdo_flag:
                    last_rdo = pkg

                payload = {
                    "timestamp": timestamp_str,
                    "time_sec": current_time,
                    "data": pkg,
                    "is_pdo": is_pdo_flag,
                    "is_rdo": is_rdo_flag,
                    "is_sink_cap": is_sink_cap_flag,
                    "last_pdo": last_pdo,
                    "last_rdo": last_rdo,
                }

                try:
                    data_queue.put_nowait(payload)
                except queue.Full:
                    pass

            except Exception as exc:
                err_text = str(exc).lower()
                if "read error" in err_text:
                    data_queue.put_nowait({"error": "device_disconnected"})
                    break
                time.sleep(0.01)

    except Exception as exc:
        data_queue.put_nowait({"error": f"connection_failed: {exc}"})
    finally:
        try:
            k2.close()
        except Exception:
            pass


# ======================== PDO/RDO 解析模块 ========================
class PDParser:
    """PDO/RDO 解析器"""

    @staticmethod
    def parse_pdo(pd_payload: Dict[str, Any]) -> List[Dict[str, str]]:
        pkg = pd_payload.get("data")
        if not pkg:
            return []

        entries = []
        try:
            is_extended = bool(pkg["Message Header"]["Extended"].value())
            data_objects = pkg[4].value() if is_extended else pkg[3].value()
        except Exception:
            return entries

        if not data_objects or not isinstance(data_objects, list):
            return entries

        for idx, obj in enumerate(data_objects):
            try:
                if getattr(obj, "value", lambda: None)() == "Empty PDO":
                    continue
            except Exception:
                pass

            summary = PDParser._safe_quick_call(obj, "quick_pdo")
            raw_bits = PDParser._safe_raw(obj)
            raw_display = PDParser._bits_to_hex(raw_bits)

            entries.append({
                "index": str(idx + 1),
                "summary": summary or "",
                "raw": raw_display,
            })

        return entries

    @staticmethod
    def parse_rdo(pd_payload: Dict[str, Any]) -> Dict[str, str]:
        pkg = pd_payload.get("data")
        info: Dict[str, str] = {}

        if pkg is None:
            info["summary"] = "Invalid RDO"
            return info

        try:
            data_objects = pkg[3].value()
        except Exception:
            info["summary"] = "Invalid RDO"
            return info

        if not data_objects or not isinstance(data_objects, list) or len(data_objects) == 0:
            info["summary"] = "Invalid RDO"
            return info

        target = data_objects[0]
        summary = PDParser._safe_quick_call(target, "quick_rdo")

        if not summary or summary == "Not a RDO":
            summary = "Invalid RDO"

        raw_bits = PDParser._safe_raw(target)
        raw_display = PDParser._bits_to_hex(raw_bits)

        try:
            position = target["Object Position"].value()
            position_text = f"Position: {position}"
        except Exception:
            position_text = None

        info["summary"] = summary
        if raw_display:
            info["raw"] = raw_display
        if position_text:
            info["details"] = position_text

        return info

    @staticmethod
    def _safe_raw(node: Any) -> str:
        try:
            return node.raw()
        except Exception:
            return ""

    @staticmethod
    def _bits_to_hex(bit_str: str) -> str:
        if not bit_str:
            return ""
        try:
            value = int(bit_str, 2)
            width = (len(bit_str) + 3) // 4
            return f"0x{value:0{width}X}"
        except Exception:
            return ""

    @staticmethod
    def _safe_quick_call(node: Any, attr: str) -> Optional[str]:
        try:
            method = getattr(node, attr)
            result = method()
            return result if isinstance(result, str) else str(result)
        except Exception:
            return None


class CableDataParser:
    """解析 Structured VDM 线缆身份信息"""

    @staticmethod
    def parse(payload: Dict[str, Any]) -> Optional[List[Tuple[str, Any]]]:
        pkg = payload.get("data") if isinstance(payload, dict) else None
        if pkg is None:
            return None

        header = CableDataParser._get_metadata(pkg, "Message Header")
        if header is None:
            return None

        msg_type = CableDataParser._get_value(header, "Message Type")
        if msg_type != "Vendor_Defined":
            return None

        sop_value = CableDataParser._get_value(pkg, "SOP*")
        if sop_value not in ("SOP'", "SOP''"):
            return None

        data_objects = CableDataParser._get_metadata(pkg, "Data Objects")
        if data_objects is None:
            return None

        vdm_entries = data_objects.value()
        if not isinstance(vdm_entries, list) or not vdm_entries:
            return None

        vdm_header = vdm_entries[0]
        if not hasattr(vdm_header, "field") or vdm_header.field() != "VDM Header":
            return None

        vdm_type = CableDataParser._get_value(vdm_header, "VDM Type")
        if vdm_type != "Structured":
            return None

        command = CableDataParser._get_value(vdm_header, "Command")
        command_type = CableDataParser._get_value(vdm_header, "Command Type")
        if command != "Discover Identity" or command_type != "ACK":
            return None

        info: List[Tuple[str, Any]] = []

        svid = CableDataParser._get_value(vdm_header, "SVID")
        if svid:
            info.append(("cable_svid", str(svid)))

        info.append(("cable_source", str(sop_value)))
        info.append(("cable_command", f"{command} ({command_type})"))

        id_header = CableDataParser._find_by_field(vdm_entries, "ID Header VDO")
        if id_header is not None:
            vid = CableDataParser._get_value(id_header, "USB Vendor ID")
            vendor_name = CableDataParser._resolve_vendor_name(vid)
            if vid:
                vid_display = str(vid).upper()
                if vendor_name:
                    vid_display = f"{vid_display} ({vendor_name})"
                info.append(("cable_vendor_id", vid_display))

            cable_role = None
            if sop_value in ("SOP'", "SOP''"):
                cable_role = CableDataParser._get_value(id_header, "Product Type (Cable Plug/VPD)")
            if cable_role:
                info.append(("cable_role", str(cable_role)))

        product_vdo = CableDataParser._find_by_field(vdm_entries, "Product VDO")
        if product_vdo is not None:
            product_id = CableDataParser._get_value(product_vdo, "USB Product ID")
            bcd_device = CableDataParser._get_value(product_vdo, "bcdDevice")
            if product_id:
                info.append(("cable_product_id", str(product_id).upper()))
            if bcd_device:
                info.append(("cable_device_version", str(bcd_device).upper()))

        passive_vdo = CableDataParser._find_by_field(vdm_entries, "Passive Cable VDO")
        active_vdo_1 = CableDataParser._find_by_field(vdm_entries, "Active Cable VDO 1")
        active_vdo_2 = CableDataParser._find_by_field(vdm_entries, "Active Cable VDO 2")
        vpd_vdo = CableDataParser._find_by_field(vdm_entries, "VPD VDO")

        if passive_vdo is not None:
            info.append(("cable_type", "cable_type_passive"))
            info.extend(CableDataParser._extract_passive_info(passive_vdo))
        elif active_vdo_1 is not None:
            info.append(("cable_type", "cable_type_active"))
            info.extend(CableDataParser._extract_active_info(active_vdo_1, active_vdo_2))
        elif vpd_vdo is not None:
            info.append(("cable_type", "cable_type_vpd"))
            info.extend(CableDataParser._extract_vpd_info(vpd_vdo))

        return info if info else None

    @staticmethod
    def _get_metadata(node: Any, field: str):
        if node is None or not hasattr(node, "__getitem__"):
            return None
        try:
            child = node[field]
        except Exception:
            return None
        return child if hasattr(child, "value") else None

    @staticmethod
    def _get_value(node: Any, field: str):
        metadata_node = CableDataParser._get_metadata(node, field)
        if metadata_node is None:
            return None
        try:
            return metadata_node.value()
        except Exception:
            return None

    @staticmethod
    def _find_by_field(items: List[Any], field_name: str):
        for item in items:
            if hasattr(item, "field") and item.field() == field_name:
                return item
        return None

    @staticmethod
    def _resolve_vendor_name(vid_value: Any) -> Optional[str]:
        if vid_value is None:
            return None
        text = str(vid_value).strip()
        candidates: List[Any] = [text, text.upper()]

        upper_text = text.upper()
        if upper_text.startswith("0X"):
            candidates.append(upper_text[2:])
        else:
            candidates.append(f"0x{upper_text}")

        try:
            vid_int = int(text, 16)
        except (ValueError, TypeError):
            vid_int = None

        if vid_int is not None:
            candidates.extend([vid_int, f"{vid_int:04X}", f"0x{vid_int:04X}"])

        for candidate in candidates:
            if candidate in VENDOR_IDS:
                return VENDOR_IDS[candidate]
        return None

    @staticmethod
    def _extract_passive_info(node: Any) -> List[Tuple[str, Any]]:
        rows: List[Tuple[str, Any]] = []
        rows.append(("cable_connector", CableDataParser._fmt(CableDataParser._get_value(
            node, "USB Type-C plug to USB Type-C/Captive (Passive Cable)"))))
        rows.append(("cable_termination", CableDataParser._fmt(CableDataParser._get_value(
            node, "Cable Termination Type (Passive Cable)"))))
        rows.append(("cable_max_voltage", CableDataParser._fmt(CableDataParser._get_value(
            node, "Maximum VBUS Voltage (Passive Cable)"))))
        rows.append(("cable_current", CableDataParser._fmt(CableDataParser._get_value(
            node, "VBUS Current Handling Capability (Passive Cable)"))))
        rows.append(("cable_highest_speed", CableDataParser._fmt(CableDataParser._get_value(
            node, "USB Highest Speed (Passive Cable)"))))
        rows.append(("cable_latency", CableDataParser._fmt(CableDataParser._get_value(
            node, "Cable Latency (Passive Cable)"))))
        rows.append(("cable_supports_epr", CableDataParser._fmt(CableDataParser._get_value(
            node, "EPR Capable (Passive Cable)"))))
        return [row for row in rows if row[1]]

    @staticmethod
    def _extract_active_info(primary: Any, secondary: Any) -> List[Tuple[str, Any]]:
        rows: List[Tuple[str, Any]] = []
        rows.append(("cable_connector", CableDataParser._fmt(CableDataParser._get_value(
            primary, "USB Type-C plug to USB Type-C/Captive"))))
        rows.append(("cable_termination", CableDataParser._fmt(CableDataParser._get_value(
            primary, "Cable Termination Type (Active Cable)"))))
        rows.append(("cable_max_voltage", CableDataParser._fmt(CableDataParser._get_value(
            primary, "Maximum VBUS Voltage (Active Cable)"))))
        rows.append(("cable_current", CableDataParser._fmt(CableDataParser._get_value(
            primary, "VBUS Current Handling Capability (Active Cable)"))))
        rows.append(("cable_highest_speed", CableDataParser._fmt(CableDataParser._get_value(
            primary, "USB Highest Speed (Active Cable)"))))
        rows.append(("cable_supports_epr", CableDataParser._fmt(CableDataParser._get_value(
            primary, "EPR Capable (Active Cable)"))))

        sbu_supported = CableDataParser._get_value(primary, "SBU Supported")
        if sbu_supported is not None:
            rows.append(("cable_sbu", CableDataParser._fmt(sbu_supported)))

        if secondary is not None:
            rows.append(("cable_max_temp", CableDataParser._fmt(CableDataParser._get_value(
                secondary, "Maximum Operating Temperature"))))
            rows.append(("cable_shutdown_temp", CableDataParser._fmt(CableDataParser._get_value(
                secondary, "Shutdown Temperature"))))
            rows.append(("cable_usb4", CableDataParser._fmt(CableDataParser._get_value(
                secondary, "USB4 Supported"))))

        return [row for row in rows if row[1]]

    @staticmethod
    def _extract_vpd_info(node: Any) -> List[Tuple[str, Any]]:
        rows: List[Tuple[str, Any]] = []
        rows.append(("cable_max_voltage", CableDataParser._fmt(CableDataParser._get_value(
            node, "Maximum VBUS Voltage"))))
        rows.append(("cable_charge_through", CableDataParser._fmt(CableDataParser._get_value(
            node, "Charge Through Support"))))
        rows.append(("cable_charge_through_current", CableDataParser._fmt(CableDataParser._get_value(
            node, "Charge Through Current Support"))))
        rows.append(("cable_vbus_impedance", CableDataParser._fmt(CableDataParser._get_value(
            node, "VBUS Impedance"))))
        rows.append(("cable_ground_impedance", CableDataParser._fmt(CableDataParser._get_value(
            node, "Ground Impedance"))))
        return [row for row in rows if row[1]]

    @staticmethod
    def _fmt(value: Any) -> Any:
        if value is None:
            return ""
        if isinstance(value, bool):
            return "bool_true" if value else "bool_false"
        return str(value)


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

        self._build_ui()
        self._populate_device_list()
        self._apply_language()

        self.poll_timer = QTimer(self)
        self.poll_timer.timeout.connect(self._poll_queue)
        self.poll_timer.start(50)

    def _text(self, key: str, **kwargs) -> str:
        language_map = LANG_STRINGS.get(self.current_language, LANG_STRINGS["zh"])
        fallback_map = LANG_STRINGS["zh"]
        text = language_map.get(key, fallback_map.get(key, key))
        if kwargs:
            try:
                text = text.format(**kwargs)
            except Exception:
                pass
        return text

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

        self._update_language_button_text()
        self._update_connect_button_text()
        self._update_start_button_text()
        self._set_status_message(self.current_status_key)
        self._update_count_label()
        self._update_table_headers()
        self._update_cable_info(self.cable_info_rows)
        self._refresh_device_selector_labels()

    def _language_display(self, code: str) -> str:
        if code in LANG_STRINGS:
            lang_map = LANG_STRINGS[code]
            display = lang_map.get(f"language_option_{code}")
            if display:
                return display
        active_map = LANG_STRINGS.get(self.current_language, LANG_STRINGS["zh"])
        return active_map.get(f"language_option_{code}", code)

    def _update_language_button_text(self) -> None:
        if not hasattr(self, "language_button"):
            return
        target_language = "en" if self.current_language == "zh" else "zh"
        display = self._language_display(target_language)
        self.language_button.setText(display)
        self.language_button.setToolTip(self._text("language_switch_tooltip", language=display))

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

    def _translate_cable_field(self, key: str) -> str:
        return CABLE_FIELD_TEXT.get(key, {}).get(self.current_language, key)

    def _translate_cable_value(self, value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, str) and value in CABLE_VALUE_TEXT:
            return CABLE_VALUE_TEXT[value].get(self.current_language, value)
        if isinstance(value, str):
            if value == "bool_true":
                return CABLE_VALUE_TEXT["bool_true"].get(self.current_language, "Yes")
            if value == "bool_false":
                return CABLE_VALUE_TEXT["bool_false"].get(self.current_language, "No")
            return value
        return str(value)

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

        self.language_button = QPushButton()
        self.language_button.setMinimumWidth(96)
        self.language_button.clicked.connect(self._on_language_button_clicked)
        control_layout.addWidget(self.language_button)

        control_layout.addStretch(1)

        self.status_label = QLabel()
        self.status_label.setStyleSheet(f"color: {STATUS_COLORS['status_disconnected']};")
        control_layout.addWidget(self.status_label)

        self.count_label = QLabel()
        control_layout.addWidget(self.count_label)

        root_layout.addLayout(control_layout)

        splitter = QSplitter(Qt.Horizontal)
        root_layout.addWidget(splitter, 1)

        cable_container = QWidget()
        cable_layout = QVBoxLayout(cable_container)
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

        cable_container.setMinimumWidth(320)
        splitter.addWidget(cable_container)

        pdo_container = QWidget()
        pdo_layout = QVBoxLayout(pdo_container)
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

        splitter.addWidget(pdo_container)

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
        header.sectionResized.connect(self._on_summary_column_resized)
        self.table.itemSelectionChanged.connect(self._on_selection_changed)
        splitter.addWidget(self.table)

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

        splitter.addWidget(detail_container)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        splitter.setStretchFactor(2, 4)
        splitter.setStretchFactor(3, 3)
        splitter.setSizes([340, 280, 520, 360])

        self.summary_metrics = QFontMetrics(self.table.font())
        if self.summary_metrics:
            self.summary_char_px = max(6, self.summary_metrics.horizontalAdvance("0"))
            self.summary_line_height = max(16, self.summary_metrics.lineSpacing())
        base_height = self.summary_line_height + self.summary_base_padding
        self.table.verticalHeader().setDefaultSectionSize(base_height)

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

        if not HID_AVAILABLE or hid is None:
            return entries

        try:
            enumerated = hid.enumerate(K2_TARGET_VID, K2_TARGET_PID)
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

    def _on_language_button_clicked(self) -> None:
        target_language = "en" if self.current_language == "zh" else "zh"
        if target_language == self.current_language or target_language not in LANG_STRINGS:
            return
        self.current_language = target_language
        self._apply_language()
        if not self.device_open:
            self._populate_device_list()
        else:
            self._refresh_device_selector_labels()

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
        self._update_connect_button_text()
        self._update_start_button_text()
        if hasattr(self, "start_btn"):
            self.start_btn.setEnabled(False)
        self._set_status_message("status_disconnected")

        if HID_AVAILABLE and hasattr(self, "refresh_devices_btn"):
            self.refresh_devices_btn.setEnabled(True)
        if HID_AVAILABLE and hasattr(self, "device_selector"):
            self.device_selector.setEnabled(True)
            self._populate_device_list()

    def _toggle_collection(self):
        if not self.device_open:
            return

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

    def _export_csv(self):
        if not self.log_records:
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
                ])

                for record in self.log_records:
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

            QMessageBox.information(self, self._text("dialog_export_success_title"), self._text("dialog_export_success", count=len(self.log_records)))
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

                    self.log_records.append(record)
                    self.log_index = max(self.log_index, index)
                    self._add_record_to_table(record)
                    imported_count += 1

                except Exception as exc:
                    print(self._text("log_skip_invalid_row", detail=exc))
                    continue

            self._update_count_label()
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

    def _handle_payload(self, payload: Dict[str, Any]):
        error_info = payload.get("error")
        if error_info:
            self._show_error(str(error_info))
            return

        if self.is_paused:
            return

        cable_rows = CableDataParser.parse(payload)
        if cable_rows:
            self._update_cable_info(cable_rows)

        if payload.get("is_pdo"):
            pdo_data = PDParser.parse_pdo(payload)
            if pdo_data is not None:
                self._update_current_pdos(pdo_data)
            if pdo_data:
                summary = " | ".join(entry["summary"] for entry in pdo_data if entry["summary"].strip())
                if summary.strip():
                    self._add_record(payload, "PDO", summary, pdo_data)

        if payload.get("is_rdo"):
            rdo_info = PDParser.parse_rdo(payload)
            summary = rdo_info.get("summary", "")
            if summary and summary != "Invalid RDO":
                self._add_record(payload, "RDO", summary, rdo_info)

    def _add_record(self, payload: Dict[str, Any], record_type: str, summary: str, data: Any):
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

        self.log_records.append(record)
        self._add_record_to_table(record)
        self._update_count_label()

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
