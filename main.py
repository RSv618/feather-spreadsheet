import sys
import os
import pyarrow.feather as feather
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import Qt, QModelIndex
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableView, QAction, QFileDialog,
    QMessageBox, QMenu, QLabel, QStatusBar, QComboBox, QDialog,
    QVBoxLayout, QHBoxLayout, QPushButton, QWidget)
from PyQt5.QtGui import QIcon

# Determine the base path
if getattr(sys, 'frozen', False):  # Running as PyInstaller executable
    BASE_PATH = sys._MEIPASS  # Temporary folder where files are extracted
else:
    BASE_PATH = os.path.dirname(__file__)  # Script directory during development

# Construct the full path to logo.ico
ICON_PATH = os.path.join(BASE_PATH, "logo.ico")

class TableModel(QtCore.QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data
        self._original_data = data.copy()
        self.undo_stack = []
        self.redo_stack = []
        self.show_dtypes = False

    def data(self, index, role):
        if not index.isValid():
            return None

        if role == Qt.DisplayRole:
            value = self._data.iloc[index.row(), index.column()]
            return str(value)

        if role == Qt.EditRole:
            value = self._data.iloc[index.row(), index.column()]
            return str(value)

        # Removed background color change based on data type

    def setData(self, index, value, role):
        if role == Qt.EditRole:
            row = index.row()
            col = index.column()
            try:
                # Store current state for undo
                self.push_undo_state()

                # Convert the string value to the appropriate type
                dtype = self._data.dtypes.iloc[col]
                col_name = self._data.columns[col]

                # Try to convert the input value to the column's data type
                if pd.api.types.is_numeric_dtype(dtype):
                    self._data.loc[self._data.index[row], col_name] = pd.to_numeric(value)
                elif pd.api.types.is_datetime64_any_dtype(dtype):
                    self._data.loc[self._data.index[row], col_name] = pd.to_datetime(value)
                else:
                    self._data.loc[self._data.index[row], col_name] = value

                # Clear redo stack when a new edit is made
                self.redo_stack = []

                self.dataChanged.emit(index, index, [role])
                return True
            except Exception as e:
                QMessageBox.critical(None, "Error", f"Cannot convert '{value}' to {dtype}: {str(e)}")
                return False
        return False

    def push_undo_state(self):
        self.undo_stack.append(self._data.copy())
        if len(self.undo_stack) > 30:  # Limit stack size
            self.undo_stack.pop(0)

    def undo(self):
        if self.undo_stack:
            # Save current state for redo
            self.redo_stack.append(self._data.copy())
            # Restore previous state
            self._data = self.undo_stack.pop()
            self.layoutChanged.emit()
            return True
        return False

    def redo(self):
        if self.redo_stack:
            # Save current state for undo
            self.undo_stack.append(self._data.copy())
            # Restore next state
            self._data = self.redo_stack.pop()
            self.layoutChanged.emit()
            return True
        return False

    def change_column_type(self, col_index, new_type):
        col_name = self._data.columns[col_index]
        try:
            # Store current state for undo
            self.push_undo_state()

            # Convert column to the new type
            if new_type == "int":
                self._data[col_name] = pd.to_numeric(self._data[col_name], errors='coerce').astype('Int64')
            elif new_type == "float":
                self._data[col_name] = pd.to_numeric(self._data[col_name], errors='coerce')
            elif new_type == "str" or new_type == "string":
                self._data[col_name] = self._data[col_name].astype(str)
            elif new_type == "datetime":
                self._data[col_name] = pd.to_datetime(self._data[col_name], errors='coerce')
            elif new_type == "bool" or new_type == "boolean":
                self._data[col_name] = self._data[col_name].astype(bool)
            elif new_type == "category":
                self._data[col_name] = self._data[col_name].astype('category')
            else:
                QMessageBox.warning(None, "Warning", f"Type '{new_type}' is not supported for conversion.")
                return False

            self.layoutChanged.emit()
            return True
        except Exception as e:
            QMessageBox.critical(None, "Error", f"Failed to convert column to {new_type}: {str(e)}")
            return False

    def flags(self, index):
        return super().flags(index) | Qt.ItemIsEditable

    def rowCount(self, index=QModelIndex()):
        return self._data.shape[0]

    def columnCount(self, index=QModelIndex()):
        return self._data.shape[1]

    def headerData(self, section, orientation, role):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                header = str(self._data.columns[section])
                if self.show_dtypes:
                    dtype = self._data.dtypes.iloc[section]
                    header += f" ({dtype})"
                return header
            if orientation == Qt.Vertical:
                return str(self._data.index[section])
        return None

    def set_dtype_visibility(self, show):
        self.show_dtypes = show
        self.headerDataChanged.emit(Qt.Horizontal, 0, self.columnCount() - 1)

    def get_data(self):
        return self._data


class ChangeDataTypeDialog(QDialog):
    def __init__(self, current_type, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Change Column Data Type")

        layout = QVBoxLayout()

        # Type selection
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("New Type:"))

        self.type_combo = QComboBox()
        self.type_combo.addItems(["int", "float", "str", "datetime", "bool", "category"])
        self.type_combo.setCurrentText(self._get_simplified_type(current_type))
        type_layout.addWidget(self.type_combo)

        layout.addLayout(type_layout)

        # Buttons
        btn_layout = QHBoxLayout()

        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        btn_layout.addWidget(ok_button)

        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_button)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def _get_simplified_type(self, dtype):
        dtype_str = str(dtype)
        if "int" in dtype_str:
            return "int"
        elif "float" in dtype_str:
            return "float"
        elif "datetime" in dtype_str:
            return "datetime"
        elif "bool" in dtype_str:
            return "bool"
        elif "category" in dtype_str:
            return "category"
        else:
            return "str"

    def get_selected_type(self):
        return self.type_combo.currentText()


class FindDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Find in Table")
        self.resize(300, 100)

        layout = QVBoxLayout()

        # Search input
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Find:"))

        self.search_text = QtWidgets.QLineEdit()
        search_layout.addWidget(self.search_text)

        layout.addLayout(search_layout)

        # Buttons
        btn_layout = QHBoxLayout()

        find_button = QPushButton("Find Next")
        find_button.clicked.connect(self.accept)
        btn_layout.addWidget(find_button)

        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_button)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def get_search_text(self):
        return self.search_text.text()


class WelcomeWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout()

        # Add open button
        open_button = QPushButton("Load feather file...")
        open_button.setStyleSheet("padding: 15px;")
        open_button.clicked.connect(parent.open_file)
        layout.addWidget(open_button, alignment=Qt.AlignCenter)

        self.setLayout(layout)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Feather Spreadsheet")
        self.resize(300, 200)

        # Set application icon if available
        if os.path.exists(ICON_PATH):
            self.setWindowIcon(QIcon(ICON_PATH))

        # Apply light mode at the start
        self.dark_mode = False
        self.apply_light_mode()

        # Create central widget container
        self.central_container = QWidget()
        self.setCentralWidget(self.central_container)
        self.container_layout = QVBoxLayout(self.central_container)

        # Create table view
        self.table = QTableView()
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.context_menu)

        # Set up initial empty data
        self.model = None

        # Set clipboard and selection mode
        self.clipboard = QApplication.clipboard()
        self.table.setSelectionMode(QTableView.ContiguousSelection)

        # Status bar
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.status_label = QLabel("")
        self.statusBar.addWidget(self.status_label)

        # Create menu bar
        self.create_menus()

        # Set up keyboard shortcuts - Fixed approach for better compatibility
        self.setup_shortcuts()

        # Check for command line arguments
        self.current_file = None
        if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
            self.load_file(sys.argv[1])
        else:
            # Show welcome widget instead of sample data
            self.show_welcome_screen()

        # Set dark mode flag
        self.dark_mode = False

    def setup_shortcuts(self):
        # Using QAction approach for better shortcut handling
        self.copy_action.setShortcut(QtGui.QKeySequence.Copy)
        self.copy_action.setShortcutContext(Qt.ApplicationShortcut)

        self.paste_action.setShortcut(QtGui.QKeySequence.Paste)
        self.paste_action.setShortcutContext(Qt.ApplicationShortcut)

        self.undo_action.setShortcut(QtGui.QKeySequence.Undo)
        self.undo_action.setShortcutContext(Qt.ApplicationShortcut)

        self.redo_action.setShortcut(QtGui.QKeySequence.Redo)
        self.redo_action.setShortcutContext(Qt.ApplicationShortcut)

        self.find_action.setShortcut(QtGui.QKeySequence.Find)
        self.find_action.setShortcutContext(Qt.ApplicationShortcut)

    def create_menus(self):
        # Create menu bar
        menubar = self.menuBar()

        # File menu
        file_menu = menubar.addMenu('&File')

        open_action = QAction('&Open', self)
        open_action.setShortcut('Ctrl+O')
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)

        save_action = QAction('&Save', self)
        save_action.setShortcut('Ctrl+S')
        save_action.triggered.connect(self.save_file)
        file_menu.addAction(save_action)

        save_as_action = QAction('Save &As', self)
        save_as_action.setShortcut('Ctrl+Shift+S')
        save_as_action.triggered.connect(self.save_file_as)
        file_menu.addAction(save_as_action)

        file_menu.addSeparator()

        exit_action = QAction('&Exit', self)
        exit_action.setShortcut('Ctrl+Q')
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Edit menu
        edit_menu = menubar.addMenu('&Edit')

        self.copy_action = QAction('&Copy', self)
        self.copy_action.triggered.connect(self.copy_selection)
        edit_menu.addAction(self.copy_action)

        self.paste_action = QAction('&Paste', self)
        self.paste_action.triggered.connect(self.paste)
        edit_menu.addAction(self.paste_action)

        edit_menu.addSeparator()

        self.undo_action = QAction('&Undo', self)
        self.undo_action.triggered.connect(self.undo)
        edit_menu.addAction(self.undo_action)

        self.redo_action = QAction('&Redo', self)
        self.redo_action.triggered.connect(self.redo)
        edit_menu.addAction(self.redo_action)

        edit_menu.addSeparator()

        self.find_action = QAction('&Find', self)
        self.find_action.triggered.connect(self.find_in_table)
        edit_menu.addAction(self.find_action)

        # View menu
        view_menu = menubar.addMenu('&View')

        self.dark_mode_action = QAction('&Dark Mode', self, checkable=True)
        self.dark_mode_action.triggered.connect(self.toggle_dark_mode)
        view_menu.addAction(self.dark_mode_action)

        self.show_dtypes_action = QAction('Show Column &Data Types', self, checkable=True)
        self.show_dtypes_action.triggered.connect(self.toggle_dtypes)
        view_menu.addAction(self.show_dtypes_action)

    def show_welcome_screen(self):
        # Clear existing widgets in container
        for i in reversed(range(self.container_layout.count())):
            self.container_layout.itemAt(i).widget().setParent(None)

        # Add welcome widget
        welcome_widget = WelcomeWidget(self)
        self.container_layout.addWidget(welcome_widget)

    def show_table(self):
        # Clear existing widgets in container
        for i in reversed(range(self.container_layout.count())):
            self.container_layout.itemAt(i).widget().setParent(None)

        # Add table widget
        self.container_layout.addWidget(self.table)

    def open_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open Feather File", "", "Feather Files (*.feather);;All Files (*)",
            options=options
        )

        if file_path:
            self.load_file(file_path)

    def load_file(self, file_path):
        try:
            data = feather.read_feather(file_path)
            self.model = TableModel(data)
            self.table.setModel(self.model)
            self.current_file = file_path
            self.setWindowTitle(f"Feather Spreadsheet - {os.path.basename(file_path)}")
            self.statusBar.showMessage(f"Loaded file: {file_path}", 5000)
            self.show_table()

            # Auto-resize columns to content
            self.table.resizeColumnsToContents()

            # Calculate appropriate window size
            screen_geo = QApplication.primaryScreen().availableGeometry()

            # Calculate width components
            vertical_header_width = self.table.verticalHeader().width()
            columns_width = sum(self.table.columnWidth(i) for i in range(self.model.columnCount()))
            total_width = vertical_header_width + columns_width + 30  # Add margins/scrollbar

            # Set reasonable width limits
            max_width = int(screen_geo.width() * 0.8)
            desired_width = min(total_width, max_width)

            # Calculate height components
            header_height = self.table.horizontalHeader().height()
            row_height = self.table.rowHeight(0) if self.model.rowCount() > 0 else 30
            visible_rows = min(20, self.model.rowCount())
            total_height = header_height + (row_height * visible_rows) + 120  # Add UI margins

            # Set reasonable height limits
            max_height = int(screen_geo.height() * 0.7)
            desired_height = min(total_height, max_height)

            # Apply new window size
            self.resize(desired_width, desired_height)
            self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open file: {str(e)}")

    def save_file(self):
        if self.model is None:
            self.statusBar.showMessage("No data to save", 3000)
            return

        if self.current_file:
            self.save_to_file(self.current_file)
        else:
            self.save_file_as()

    def save_file_as(self):
        if self.model is None:
            self.statusBar.showMessage("No data to save", 3000)
            return

        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Feather File", "", "Feather Files (*.feather);;All Files (*)",
            options=options
        )

        if file_path:
            self.save_to_file(file_path)

    def save_to_file(self, file_path):
        try:
            data = self.model.get_data()
            feather.write_feather(data, file_path)
            self.current_file = file_path
            self.setWindowTitle(f"Feather Spreadsheet - {os.path.basename(file_path)}")
            self.statusBar.showMessage(f"Saved to: {file_path}", 5000)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save file: {str(e)}")

    def copy_selection(self):
        if self.model is None:
            return

        selection = self.table.selectionModel()
        if not selection.hasSelection():
            return

        selected_indexes = selection.selectedIndexes()
        if not selected_indexes:
            return

        # Get row and column boundaries
        rows = set()
        cols = set()
        for idx in selected_indexes:
            rows.add(idx.row())
            cols.add(idx.column())

        min_row, max_row = min(rows), max(rows)
        min_col, max_col = min(cols), max(cols)

        # Create a string representation of the selected cells
        selection_text = []
        for row in range(min_row, max_row + 1):
            row_data = []
            for col in range(min_col, max_col + 1):
                index = self.model.createIndex(row, col)
                cell_data = self.model.data(index, Qt.DisplayRole)
                row_data.append(str(cell_data))
            selection_text.append('\t'.join(row_data))

        # Copy to clipboard
        self.clipboard.setText('\n'.join(selection_text))
        self.statusBar.showMessage(f"Copied selection: {min_row}:{max_row}, {min_col}:{max_col}", 3000)

    def paste(self):
        if self.model is None:
            return

        selection = self.table.selectionModel()
        if not selection.hasSelection():
            return

        # Get current cursor position
        current = selection.currentIndex()
        if not current.isValid():
            return

        start_row = current.row()
        start_col = current.column()

        # Get clipboard text
        clipboard_text = self.clipboard.text()
        if not clipboard_text:
            return

        # Parse clipboard data
        rows = clipboard_text.split('\n')
        for r, row_text in enumerate(rows):
            if not row_text:
                continue

            cells = row_text.split('\t')
            for c, cell_text in enumerate(cells):
                row = start_row + r
                col = start_col + c

                if row < self.model.rowCount() and col < self.model.columnCount():
                    index = self.model.createIndex(row, col)
                    self.model.setData(index, cell_text, Qt.EditRole)

    def undo(self):
        if self.model is None:
            return

        if self.model.undo():
            self.statusBar.showMessage("Undo successful", 3000)
        else:
            self.statusBar.showMessage("Nothing to undo", 3000)

    def redo(self):
        if self.model is None:
            return

        if self.model.redo():
            self.statusBar.showMessage("Redo successful", 3000)
        else:
            self.statusBar.showMessage("Nothing to redo", 3000)

    def apply_light_mode(self):
        # Create and apply light mode style sheet with slightly darker headers
        light_style = """
        QTableView {
            border: 1px solid #B2B2B2;
        }
        QHeaderView::section {
            background-color: #E6E6E6;
            border-bottom: 1px solid #B2B2B2;
            border-top: none;
            border-right: 1px solid #B2B2B2;
            border-left: none;
        }
        QTableView QTableCornerButton::section {
            background-color: #E6E6E6;
            border-top: none;
            border-left: none;
            border-right: 1px solid #B2B2B2;
            border-bottom: 1px solid #B2B2B2;
        }
        """
        self.setStyleSheet(light_style)

    def toggle_dark_mode(self):
        self.dark_mode = self.dark_mode_action.isChecked()

        if self.dark_mode:
            # Create and apply dark style sheet
            dark_style = """
            QMainWindow, QTableView, QHeaderView, QDialog, QWidget {
                background-color: #2d2d2d;
                color: #e0e0e0;
            }
            QTableView {
                gridline-color: #5a5a5a;
                selection-background-color: #3a3a5a;
                selection-color: white;
            }
            QTableView QTableCornerButton::section {
                background-color: #3a3a3a;
            }
            QHeaderView::section {
                background-color: #3a3a3a;
                color: #e0e0e0;
                border: 1px solid #5a5a5a;
            }
            QMenuBar, QMenu {
                background-color: #2d2d2d;
                color: #e0e0e0;
            }
            QMenuBar::item:selected, QMenu::item:selected {
                background-color: #3a3a5a;
            }
            QStatusBar {
                background-color: #3a3a3a;
                color: #e0e0e0;
            }
            QLabel, QPushButton, QComboBox {
                color: #e0e0e0;
                background-color: #3a3a3a;
                border: 1px solid #5a5a5a;
            }
            QPushButton:hover {
                background-color: #4a4a6a;
            }
            """
            self.setStyleSheet(dark_style)
        else:
            # Reset to default style
            self.apply_light_mode()

    def toggle_dtypes(self):
        if self.model is None:
            return

        show_dtypes = self.show_dtypes_action.isChecked()
        self.model.set_dtype_visibility(show_dtypes)

    def context_menu(self, position):
        if self.model is None:
            return

        # Get the index under cursor
        index = self.table.indexAt(position)

        # Create context menu
        context_menu = QMenu(self)

        # Add copy action
        copy_action = context_menu.addAction("Copy")
        copy_action.triggered.connect(self.copy_selection)

        # Add paste action
        paste_action = context_menu.addAction("Paste")
        paste_action.triggered.connect(self.paste)

        context_menu.addSeparator()

        # Add column-specific actions if clicked on a valid cell
        if index.isValid():
            col_index = index.column()
            col_name = self.model._data.columns[col_index]
            col_type = self.model._data.dtypes.iloc[col_index]

            change_type_action = context_menu.addAction(f"Change column '{col_name}' type...")
            change_type_action.triggered.connect(lambda: self.change_column_type(col_index))

        # Show the menu
        context_menu.exec_(self.table.viewport().mapToGlobal(position))

    def change_column_type(self, col_index):
        current_type = self.model._data.dtypes.iloc[col_index]
        dialog = ChangeDataTypeDialog(current_type, self)

        if dialog.exec_():
            new_type = dialog.get_selected_type()
            if self.model.change_column_type(col_index, new_type):
                self.statusBar.showMessage(f"Column type changed to {new_type}", 3000)

    def find_in_table(self):
        if self.model is None:
            return

        dialog = FindDialog(self)
        self.last_find_row = -1
        self.last_find_col = -1

        if dialog.exec_():
            self.find_text(dialog.get_search_text())

    def find_text(self, text):
        if not text:
            return

        # Start searching from current selection or from beginning
        selection = self.table.selectionModel().selectedIndexes()
        if selection:
            start_row = selection[-1].row()
            start_col = selection[-1].column()
        else:
            start_row = self.last_find_row
            start_col = self.last_find_col

        # Search pattern
        found = False
        for col in range(self.model.columnCount()):
            for row in range(self.model.rowCount()):
                # Ensure we start after our last position and wrap around
                if (row > start_row) or (row == start_row and col > start_col) or (
                        start_row == self.model.rowCount() - 1 and start_col == self.model.columnCount() - 1):
                    index = self.model.createIndex(row, col)
                    cell_data = self.model.data(index, Qt.DisplayRole)
                    if cell_data and text.lower() in str(cell_data).lower():
                        # Found it - select the cell
                        self.table.setCurrentIndex(index)
                        self.table.scrollTo(index)
                        self.last_find_row = row
                        self.last_find_col = col
                        return

        # If we got here, no matches were found after the current position
        # Try from the beginning if we didn't start from beginning
        if start_row > 0 or start_col > 0:
            self.last_find_row = -1
            self.last_find_col = -1
            self.find_text(text)
        else:
            self.statusBar.showMessage(f"Text '{text}' not found", 3000)

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Set application icon
    if os.path.exists(ICON_PATH):
        app.setWindowIcon(QIcon(ICON_PATH))

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())