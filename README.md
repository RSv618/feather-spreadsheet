
# Feather File Spreadsheet

A PyQt5-based application for viewing and editing Feather files with data manipulation capabilities.
## Features

- **View and Edit Feather Files**: Open, modify, and save Apache Feather files
- **Data Type Support**: Handle numeric, datetime, string, boolean, and categorical data types
- **Undo/Redo Functionality**: 30-step undo/redo capability for data edits
- **Dark/Light Mode**: Toggle between dark and light themes
- **Column Type Conversion**: Change data types through context menu
- **Search Functionality**: Find text within the table
- **Clipboard Support**: Copy/paste data selections
- **Data Type Visibility**: Show/hide column data types in headers
- **Context Menu**: Right-click for quick column operations


## Installation

1. Requirements:
- Python 3.7+
- Required Package:
```bash
pip install pyqt5 pyarrow pandas pyinstaller
```
2. Run from Source:
```bash
python main.py [filename.feather]  # Optional file argument
```
## Usage/Examples

- File Operations:
    - Ctrl+O: Open Feather file
    - Ctrl+S: Save
    - Ctrl+Shift+S: Save As
- Editing:
    - Ctrl+C: Copy selection
    - Ctrl+V: Paste
    - Ctrl+Z: Undo
    - Ctrl+Y: Redo
    - Ctrl+F: Find in table
- View Options:
    - Toggle dark mode from View menu
    - Show/hide column data types from View menu
- Column Operations:
    - Right-click column header for type conversion
    - Supported types: int, float, string, datetime, bool, category


## Building the executable

To create a standalone executable:
```bash
pyinstaller --onefile --noconsole --clean --icon "logo.ico" --add-data "logo.ico;." main.py
```
Notes:
- Ensure logo.ico exists in your working directory
- The generated executable will be in the dist/ folder
- Tested with PyInstaller 6.0+