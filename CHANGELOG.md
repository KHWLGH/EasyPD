# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.0]

### Added
- **Data Visualization**: Added real-time charts for Voltage, Current, and Power monitoring.
- **Auto Pause**: Implemented automatic data collection pause based on configurable Voltage/Current thresholds and delay settings.
- **Status Bar**: Added real-time display of Voltage, Current, and Power measurements in the status bar.
- **Performance**: Implemented batch UI updates (`_batch_update_ui`) to reduce UI freezing during high-frequency data capture.
- **Memory Management**: Added memory usage monitoring.

### Changed
- **Refactoring**: Split the monolithic `EasyPD.py` into modular components:
    - `device_comm.py`: Device communication logic.
    - `i18n.py`: Internationalization strings.
    - `pd_decoder.py`: Protocol decoding logic.
- **UI Layout**: Improved main window layout with resizable splitters for better visibility of different data sections.
- **Export**: Updated CSV export to include Voltage, Current, and Power measurement data.

## [0.1.0] - 2025-11-09

### Added
- Initial release of EasyPD.
- Real-time USB-PD protocol monitoring and data capture.
- PDO (Power Data Object) and RDO (Request Data Object) parsing.
- Cable information (VDM) identification and parsing.
- CSV data export and import functionality.
- Multi-language support (English/Chinese).
- Dark theme UI.
- Data collection pause/resume support.
