[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)

# EasyPD

[English](#english) | [中文](#中文) | [Changelog](CHANGELOG.md)

---

## 中文

### 简介

> **⚠️ 状态通知：本项目即将归档，将停止开发和活跃维护。**

**注意：本软件大部分使用Copilot等VibeCoding工具制作，可能存在未知问题**

EasyPD 是一款开源的 USB Power Delivery (USB-PD) 协议分析工具，提供直观的图形界面用于实时捕获、分析和记录 USB-PD 通信数据。支持 PDO（Power Data Object）和 RDO（Request Data Object）解析，以及线缆信息（VDM）识别。

查看 [更新日志](CHANGELOG.md) 了解版本更新详情。

### 主要特性

- 🔌 **实时数据捕获** - 实时监控 USB-PD 通信协议
- 📊 **PDO/RDO 解析** - 自动解析电源能力对象和请求对象
- 🔍 **线缆信息识别** - 支持 VDM（Vendor Defined Message）解析，显示线缆详细参数
- 📁 **数据导出/导入** - 支持 CSV 格式数据导出和导入
- 🌍 **多语言支持** - 内置中文和英文界面切换
- 🎨 **深色主题** - 舒适的深色界面设计
- ⏸️ **暂停/恢复** - 支持数据采集暂停和恢复功能
- 📝 **详细日志** - 完整的时间戳和相对时间记录

### 系统要求

- Windows 10 或更高版本
- Python 3.7+（开发运行）
- 兼容的 USB-PD 采集硬件（如 WITRN K2）

### 安装

#### 从源码运行

1. 克隆仓库：
```bash
git clone https://github.com/KHWLGH/EasyPD.git
cd EasyPD
```

2. 安装依赖：
```bash
pip install -r requirements.txt
```

3. 运行程序：
```bash
python EasyPD.py
```

#### 使用编译版本

从 [Releases](https://github.com/KHWLGH/EasyPD/releases) 页面下载最新的可执行文件，直接运行即可。

### 依赖项

- `PySide6` - Qt 6 的 Python 绑定
- `witrnhid` - WITRN 设备通信库
- `hidapi` - HID 设备访问库

详细依赖列表请参阅 [`requirements.txt`](requirements.txt)

### 使用说明

1. **连接设备**
   - 启动程序后，从下拉菜单中选择你的 USB-PD 采集设备
   - 点击"连接设备"按钮

2. **开始采集**
   - 点击"开始收集"开始捕获数据
   - 使用"暂停收集"/"继续收集"控制数据采集

3. **查看数据**
   - 主表格显示所有捕获的 PDO/RDO 记录
   - 左侧面板显示当前 PDO 列表和线缆信息
   - 点击记录查看详细信息

4. **导出数据**
   - 点击"导出 CSV"将记录保存为 CSV 文件
   - 支持导入之前保存的 CSV 文件

### 编译为可执行文件

使用 Nuitka 编译：

```bash
nuitka --onefile --windows-console-mode=disable --windows-icon-from-ico='./favicon.ico' --enable-plugin=pyside6 EasyPD.py
```

### 项目结构

```
EasyPD/
├── EasyPD.py              # 主程序（入口）
├── device_comm.py        # 设备通信封装（witrnhid / hid 接口）
├── pd_decoder.py         # USB-PD 数据解析逻辑
├── i18n.py               # 国际化字符串
├── vendor_ids_dict.py    # 厂商 ID 字典
├── requirements.txt      # 运行依赖
├── CHANGELOG.md          # 更新日志
├── README.md             # 本文件
├── LICENSE               # 许可证（GPL v3）
└── favicon.ico           # 应用图标（可选）

注：项目中还包含若干由构建或运行时生成并通常被忽略的目录/文件，例如 `.venv/`, `__pycache__/`, `EasyPD.build/`, `EasyPD.dist/`, `EasyPD.onefile-build/` 等（见 `.gitignore`）。这些不列入上面的清单以避免与版本控制中的实际源文件混淆。
```


### 贡献

欢迎贡献代码！请遵循以下步骤：

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

### 许可证

本项目采用 GNU General Public License v3.0 许可证 - 详见 [LICENSE](LICENSE) 文件

### 致谢

- 感谢 WITRN 提供的 USB-PD 采集硬件支持
- 感谢所有开源项目贡献者
- 感谢[JohnScotttt](https://github.com/JohnScotttt)

### 联系方式

- 项目主页: https://github.com/KHWLGH/EasyPD
- 问题反馈: https://github.com/KHWLGH/EasyPD/issues

---

## English

### Introduction

> **⚠️ Status Notice: This project will be archived soon and will no longer be actively maintained.**

**Note: Most of this software was created using VibeCoding tools such as Copilot, and there may be unknown issues.**

EasyPD is an open-source USB Power Delivery (USB-PD) protocol analyzer tool that provides an intuitive graphical interface for real-time capture, analysis, and recording of USB-PD communication data. It supports PDO (Power Data Object) and RDO (Request Data Object) parsing, as well as cable information (VDM) identification.

See [Changelog](CHANGELOG.md) for version history.

### Key Features

- 🔌 **Real-time Data Capture** - Monitor USB-PD communication protocol in real-time
- 📊 **PDO/RDO Parsing** - Automatically parse Power Data Objects and Request Data Objects
- 🔍 **Cable Information Recognition** - Support VDM (Vendor Defined Message) parsing with detailed cable parameters
- 📁 **Data Export/Import** - Support CSV format data export and import
- 🌍 **Multi-language Support** - Built-in Chinese and English interface switching
- 🎨 **Dark Theme** - Comfortable dark interface design
- ⏸️ **Pause/Resume** - Support data collection pause and resume functions
- 📝 **Detailed Logging** - Complete timestamp and relative time recording

### System Requirements

- Windows 10 +
- Python 3.7+ (for development)
- Compatible USB-PD capture hardware (e.g., WITRN K2)

### Installation

#### Running from Source

1. Clone the repository:
```bash
git clone https://github.com/KHWLGH/EasyPD.git
cd EasyPD
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the program:
```bash
python EasyPD.py
```

#### Using Compiled Version

Download the latest executable from the [Releases](https://github.com/KHWLGH/EasyPD/releases) page and run it directly.

### Dependencies

- `PySide6` - Python bindings for Qt 6
- `witrnhid` - WITRN device communication library
- `hidapi` - HID device access library

See [`requirements.txt`](requirements.txt) for the complete list of dependencies.

### Usage

1. **Connect Device**
   - After launching the program, select your USB-PD capture device from the dropdown menu
   - Click the "Connect" button

2. **Start Capture**
   - Click "Start Capture" to begin capturing data
   - Use "Pause Capture"/"Resume Capture" to control data collection

3. **View Data**
   - Main table displays all captured PDO/RDO records
   - Left panel shows current PDO list and cable information
   - Click on records to view detailed information

4. **Export Data**
   - Click "Export CSV" to save records as a CSV file
   - Supports importing previously saved CSV files

### Building Executable

Compile using Nuitka:

```bash
nuitka --onefile ^
       --windows-console-mode=disable ^
       --windows-icon-from-ico=favicon.ico ^
       --enable-plugin=pyside6 ^
       EasyPD.py
```

### Project Structure

```
EasyPD/
├── EasyPD.py              # Main program (entry)
├── device_comm.py        # Device communication wrapper (witrnhid / hid)
├── pd_decoder.py         # USB-PD decoding logic
├── i18n.py               # Internationalization strings
├── vendor_ids_dict.py    # Vendor ID dictionary
├── requirements.txt      # Runtime dependencies
├── CHANGELOG.md          # Changelog
├── README.md             # This file
├── LICENSE               # License (GPL v3)
└── favicon.ico           # Application icon (optional)

Note: The project also generates several build/runtime files and directories that are normally ignored by version control (see `.gitignore`), such as `.venv/`, `__pycache__/`, `EasyPD.build/`, `EasyPD.dist/`, and `EasyPD.onefile-build/`. These are omitted from the list above to avoid confusion with source files.
```

### Contributing

Contributions are welcome! Please follow these steps:

1. Fork this repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.

### Acknowledgments

- Thanks to WITRN for USB-PD capture hardware support
- Thanks to all open-source project contributors
- Thanks to [JohnScotttt](https://github.com/JohnScotttt)  

### Contact

- Project Homepage: https://github.com/KHWLGH/EasyPD
- Issue Tracker: https://github.com/KHWLGH/EasyPD/issues
