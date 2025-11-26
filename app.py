import sys
import os
import shutil
from datetime import datetime
from typing import List, TypedDict

# --- Qt Imports ---
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTableWidget, QTableWidgetItem,
    QGridLayout, QMessageBox, QLineEdit, QFileDialog, QComboBox,
    QLabel, QPushButton, QStyledItemDelegate, QProgressBar, QHBoxLayout,
    QVBoxLayout, QAbstractItemView, QStyle, QDialog, QTextBrowser, QMenuBar
)
from PyQt6.QtCore import QFile, QIODevice, Qt, QModelIndex, QSignalBlocker
from PyQt6.QtGui import QColor, QFont, QStandardItemModel, QStandardItem, QPainter, QPalette, QIcon, QAction # Added QPalette, QIcon, QAction
from PyQt6.QtCore import QTimer # Added for potential splash screen or delayed init check

# --- Data Handling Imports ---
import pandas as pd
import numpy # Often needed by pandas/pyinstaller

# --- pyexcel for .xls ---
import pyexcel as pe
# Check for the specific xls plugin
import pyexcel_xls
PYEXCEL_AVAILABLE = True

# --- openpyxl for .xlsx ---
import openpyxl
OPENPYXL_AVAILABLE = True


# --- Custom Calculation Module ---
# Assuming calc.py is in the same directory or accessible via PYTHONPATH

import calc


# === GeoMag Initialization ===
script_dir = os.path.dirname(__file__) if "__file__" in locals() else os.getcwd()
coeff_path = os.path.join(script_dir, "coeff.csv")

# Check in executable bundle path (_MEIPASS) if running frozen
if not os.path.exists(coeff_path) and getattr(sys, 'frozen', False):
    try:
        meipass_path = getattr(sys, '_MEIPASS', script_dir)
        coeff_path_frozen = os.path.join(meipass_path, "coeff.csv")
        if os.path.exists(coeff_path_frozen):
            coeff_path = coeff_path_frozen # Use the bundled path
    except Exception as e:
        print(f"Error checking MEIPASS path for coeff.csv: {e}", file=sys.stderr)

gm = None
try:
    if os.path.exists(coeff_path):
        gm = calc.GeoMag(coeff_path)
        print(f"GeoMag initialized using coefficients from: {coeff_path}")
    else:
        print(f"Warning: coeff.csv not found at '{coeff_path}'. Attempting default GeoMag init.", file=sys.stderr)
        gm = calc.GeoMag() # Attempt default if possible
except Exception as e:
    print(f"CRITICAL ERROR: Could not initialize GeoMag. Error: {e}", file=sys.stderr)
    # Error message will be shown later in the UI if possible


# === TypedDict Definitions ===
class IGRFWidgetColumnNames(TypedDict):
    latitude: str
    longitude: str
    altitude: str
    date: str

class ColourMap(TypedDict):
    colour: QColor
    string: str


# === Helper Functions ===

def parse_date(date_str):
    """Attempt to parse a date string using multiple common formats."""
    if isinstance(date_str, datetime):
        return date_str
    if isinstance(date_str, pd.Timestamp):
        return date_str.to_pydatetime()

    if not isinstance(date_str, str):
        date_str = str(date_str)
    date_str = date_str.strip()
    if not date_str:
        raise ValueError("Cannot parse empty date string.")

    formats = [
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", # ISO with T
        "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%m-%d-%Y", "%d-%m-%Y",
        "%Y%m%d", "%b %d %Y", "%B %d, %Y",
        "%a, %b %d, %Y", "%A, %B %d, %Y",
        # Add formats with time
        "%m/%d/%Y %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M",
        "%m/%d/%Y %H:%M", "%d/%m/%Y %H:%M",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except (ValueError, TypeError):
            continue

    try:
        dt_obj = pd.to_datetime(date_str, infer_datetime_format=True)
        return dt_obj.to_pydatetime()
    except (ValueError, TypeError, OverflowError) as e:
         raise ValueError(f"Date format not recognized or invalid date for '{date_str}': {e}")


def calculate_igrf_single(latitude: float, longitude: float, altitude: float, date_str: str):
    """
    Calculate IGRF for a single point. Altitude assumed in METERS.
    Returns dict with IGRF, INC, DEC, or dict with None values on error.
    """
    global gm
    if gm is None:
        print("Error: GeoMag object not initialized.", file=sys.stderr)
        return {"IGRF": None, "INC": None, "DEC": None}

    try:
        lat_f = float(latitude)
        lon_f = float(longitude)
        alt_m = float(altitude) # Altitude in Meters

        if not (-90 <= lat_f <= 90): raise ValueError(f"Latitude {lat_f} out of range (-90 to 90).")
        if not (-180 <= lon_f <= 180):
             if lon_f > 180 or lon_f < -180:
                 print(f"Warning: Longitude {lon_f} outside (-180 to 180). Adjusting.", file=sys.stderr)
                 lon_f = (lon_f + 180) % 360 - 180 # Normalize longitude

        # *** CHANGE HERE: Convert altitude from METERS to kilometers ***
        altitude_km = alt_m / 1000.0
        time_obj = parse_date(date_str) # Can raise ValueError

        # Call the GeoMag library (expects altitude in km)
        r = gm.GeoMag(time=time_obj.date(), dlat=lat_f, dlon=lon_f, h=altitude_km)

        return {"IGRF": r.ti, "INC": r.dip, "DEC": r.dec}

    except ValueError as ve:
        print(f"Data Conversion/Validation Error: {ve}. Input: Lat={latitude}, Lon={longitude}, Alt={altitude}, Date='{date_str}'", file=sys.stderr)
    except TypeError as te:
        print(f"Type Error during Calculation: {te}. Input: Lat={latitude}, Lon={longitude}, Alt={altitude}, Date='{date_str}'", file=sys.stderr)
    except Exception as e:
        print(f"Unexpected Calculation Error ({type(e).__name__}): {e}. Input: Lat={latitude}, Lon={longitude}, Alt={altitude}, Date='{date_str}'", file=sys.stderr)

    return {"IGRF": None, "INC": None, "DEC": None} # Return None dict on any error


def hex_to_colour(incolour: str) -> QColor:
    """Convert a hexadecimal color string (e.g., '#FF556C' or 'FF556C') to QColor."""
    incolour = str(incolour).lstrip('#').strip()
    if len(incolour) == 3: incolour = "".join([c*2 for c in incolour]) # Expand shorthand hex

    if len(incolour) == 6:
        try:
            return QColor(f"#{incolour}")
        except ValueError: pass
    print(f"Warning: Invalid hex color format '{incolour}', using white.", file=sys.stderr)
    return QColor(Qt.GlobalColor.white)


# === Custom Qt Delegate ===
class ColorDelegate(QStyledItemDelegate):
    """Delegate to draw combo box items with background colors and contrasting text."""
    def paint(self, painter: QPainter, option, index: QModelIndex):
        text = index.data(Qt.ItemDataRole.DisplayRole)
        bg_data = index.data(Qt.ItemDataRole.BackgroundRole)

        bg_color = QColor(Qt.GlobalColor.white) # Default
        if isinstance(bg_data, QColor): bg_color = bg_data
        elif isinstance(bg_data, str): bg_color = hex_to_colour(bg_data)

        painter.save()
        # Let Qt handle the basic background (selection, focus etc.)
        self.initStyleOption(option, index)
        style = option.widget.style() if option.widget else QApplication.style()
        style.drawPrimitive(QStyle.PrimitiveElement.PE_PanelItemViewItem, option, painter, option.widget)

        # Overlay our custom color slightly inset? Or full? Let's try full for now.
        painter.fillRect(option.rect, bg_color)

        # Text color based on contrast with our background
        luminance = (0.299 * bg_color.red() + 0.587 * bg_color.green() + 0.114 * bg_color.blue()) / 255
        text_color = QColor(Qt.GlobalColor.black) if luminance > 0.5 else QColor(Qt.GlobalColor.white)
        painter.setPen(text_color)

        text_rect = option.rect.adjusted(5, 0, -5, 0) # Padding
        elided_text = painter.fontMetrics().elidedText(text, Qt.TextElideMode.ElideRight, text_rect.width())
        painter.drawText(text_rect, Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, elided_text)
        painter.restore()


# === Custom Table Widget ===
class Table(QTableWidget):
    """Custom QTableWidget to load and display CSV/Excel data."""
    df: pd.DataFrame | None = None
    color_map: List[ColourMap] = []
    pth: str = ""

    def __init__(self, parent=None):
        super().__init__(parent)
        parent_widget = self.parentWidget()
        self.cell_font = QFont(parent_widget.font() if parent_widget else QFont())
        self.header_font = QFont(self.cell_font)
        self.header_font.setBold(True)
        self._setup_ui()

    def _setup_ui(self):
        self.setAlternatingRowColors(True)
        self.setSortingEnabled(True)
        self.setWordWrap(False)
        self.horizontalHeader().setStretchLastSection(True)
        self.verticalHeader().setVisible(False)
        self.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)

    def load_data(self, filename: str | None = None) -> bool:
        """Loads CSV, XLS, or XLSX file. Uses pyexcel for .xls."""
        actual_filename = self._get_filename(filename)
        if not actual_filename: return False
        if not self._check_file(actual_filename): return False

        self.pth = actual_filename
        self.color_map = []
        file_lower = actual_filename.lower()

        try:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            self.df = None # Reset dataframe before loading

            if file_lower.endswith('.xls'):
                if not PYEXCEL_AVAILABLE:
                    QMessageBox.critical(self, "Missing Library", "Reading .xls files requires 'pyexcel' and 'pyexcel-xls'.\nPlease install them (pip install pyexcel pyexcel-xls) and restart.")
                    return False
                print(f"Info: Reading '{os.path.basename(actual_filename)}' using pyexcel-xls.")
                import pyexcel as pe
                data = pe.get_array(file_name=actual_filename)
                if not data or len(data) < 1:
                    QMessageBox.information(self, "Empty File", f"The .xls file '{os.path.basename(actual_filename)}' appears empty or could not be read.")
                    self._clear_table(); return False
                header = data[0]
                body = data[1:]
                # Convert header items to strings robustly
                header = [str(h) if h is not None else f"Column_{i+1}" for i, h in enumerate(header)]
                self.df = pd.DataFrame(body, columns=header)

            elif file_lower.endswith('.csv'):
                print(f"Info: Reading '{os.path.basename(actual_filename)}' using pandas CSV.")
                try:
                    self.df = pd.read_csv(actual_filename, low_memory=False)
                except UnicodeDecodeError:
                    print("Info: UTF-8 failed, trying latin-1.", file=sys.stderr)
                    self.df = pd.read_csv(actual_filename, low_memory=False, encoding='latin-1')

            elif file_lower.endswith('.xlsx'):
                if not OPENPYXL_AVAILABLE:
                    QMessageBox.critical(self, "Missing Library", "Reading .xlsx files requires 'openpyxl'.\nPlease install it (pip install openpyxl) and restart.")
                    return False
                print(f"Info: Reading '{os.path.basename(actual_filename)}' using pandas/openpyxl.")
                self.df = pd.read_excel(actual_filename, engine='openpyxl')

            else: # Should be caught by _get_filename filter, but safety check
                QMessageBox.critical(self, "Unsupported Format", "Only .csv, .xls, .xlsx files are supported.")
                return False

            # --- Post-Reading Checks ---
            if self.df is None:
                 QMessageBox.critical(self, "Loading Error", "DataFrame could not be created after attempting to read the file.")
                 self._clear_table(); return False
            if self.df.empty:
                QMessageBox.information(self, "Empty DataFrame", f"The file '{os.path.basename(actual_filename)}' resulted in an empty dataset.")
                # Keep header info if present, but populate will handle empty rows
                # Let populate handle this case, just return True if headers exist?
                # Or return False? Let's return False for consistency.
                self._clear_table(); return False

            self._populate_table()
            return True

        except FileNotFoundError:
            QMessageBox.critical(self, "File Not Found", f"The file was not found:\n{actual_filename}")
        except (pd.errors.ParserError, SyntaxError) as pe: # Catch pandas/pyexcel parsing errors
             QMessageBox.critical(self, "Parsing Error", f"Could not parse the file:\n{actual_filename}\n\nError: {pe}\n\nPlease ensure valid formatting.")
        except ImportError as ie:
             lib_name = str(ie).split("'")[-2]
             QMessageBox.critical(self, "Missing Dependency", f"Reading this file requires '{lib_name}'.\nInstall it (pip install {lib_name}) and restart.")
        except Exception as e:
            QMessageBox.critical(self, "Loading Error", f"Unexpected error loading file:\n{actual_filename}\n\nError ({type(e).__name__}): {e}")
            print(f"Traceback for loading error: {e}", file=sys.stderr)

            self._clear_table(); return False # Return False on any error
        finally:
            QApplication.restoreOverrideCursor()

    def _get_filename(self, filename: str | None) -> str | None:
        if filename: return filename
        start_dir = os.path.dirname(self.pth) if self.pth else os.path.expanduser("~")
        filt = "Data files (*.csv *.xls *.xlsx);;CSV (*.csv);;Excel (*.xlsx);;Old Excel (*.xls);;All files (*)"
        diag = QFileDialog(self, "Open Data File", start_dir, filt)
        diag.setFileMode(QFileDialog.FileMode.ExistingFile)
        diag.setViewMode(QFileDialog.ViewMode.Detail)
        if diag.exec():
            selected = diag.selectedFiles()
            if selected:
                fname_lower = selected[0].lower()
                if not (fname_lower.endswith('.csv') or fname_lower.endswith('.xls') or fname_lower.endswith('.xlsx')):
                     QMessageBox.warning(self, "Unsupported Type", "Selected file must be .csv, .xls, or .xlsx.")
                     return None
                return selected[0]
        return None

    def _check_file(self, filename: str) -> bool:
        if not os.path.exists(filename):
             QMessageBox.warning(self, "File Not Found", f"File does not exist:\n{filename}"); return False
        try:
            with open(filename, 'rb') as f: f.read(1)
        except IOError as e:
            QMessageBox.critical(self, "Permission Error", f"Cannot read file (permissions/locked?):\n{filename}\nError: {e}"); return False
        return True

    def _clear_table(self):
        blocker = QSignalBlocker(self)
        try: self.setRowCount(0); self.setColumnCount(0)
        finally: pass
        self.df = None; self.pth = ""; self.color_map = []

    def _populate_table(self):
        if self.df is None: return
        self.setRowCount(0); self.setColumnCount(0) # Clear efficiently

        num_rows, num_cols = self.df.shape
        self.setRowCount(num_rows)
        self.setColumnCount(num_cols)

        headers = [str(col) if col is not None else f"Column_{j+1}" for j, col in enumerate(self.df.columns)]
        self.df.columns = headers # Ensure consistent string columns in df
        self.setHorizontalHeaderLabels(headers)
        self._style_headers(headers)

        blocker = QSignalBlocker(self)
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor) # Indicate population busy
        try:
            # Consider df.astype(str).values.tolist() if performance is critical and memory allows
            for i, row in enumerate(self.df.itertuples(index=False, name=None)):
                for j, cell_value in enumerate(row):
                    item = QTableWidgetItem("" if pd.isna(cell_value) else str(cell_value))
                    item.setFont(self.cell_font)
                    self.setItem(i, j, item)
                # Update UI occasionally during large population (optional)
                # if i % 500 == 0: QApplication.processEvents()
        finally:
            QApplication.restoreOverrideCursor() # Restore cursor

        self.resizeColumnsToContents()

    def _style_headers(self, headers: List[str]):
        """Applies colors and font to headers, populates color_map."""
        catppuccin_colors = [ # Catppuccin Mocha
            "F5E0DC", "F2CDCD", "F5C2E7", "E8A2AF", "F38BA8", "FAB387", "F9E2AF",
            "A6E3A1", "94E2D5", "89DCEB", "74C7EC", "89B4FA", "B4Befe", "CBA6F7",
            "CDD6F4", "BAC2DE", "A6ADC8", "9399B2", "7F849C", "6C7086"]
        num_colors = len(catppuccin_colors)
        self.color_map = []

        for index, header in enumerate(headers):
            item = QTableWidgetItem(header)
            item.setFont(self.header_font)
            color = hex_to_colour(catppuccin_colors[index % num_colors])
            self.color_map.append(ColourMap(string=header, colour=color))
            item.setBackground(color)
            luminance = (0.299*color.red() + 0.587*color.green() + 0.114*color.blue()) / 255
            item.setForeground(QColor(Qt.GlobalColor.black if luminance > 0.5 else Qt.GlobalColor.white))
            self.setHorizontalHeaderItem(index, item)

# === IGRF Column Selection Widget ===
class IGRFWidget(QWidget):
    """Widget for selecting Latitude, Longitude, Altitude, Date columns."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self._column_map_data: List[ColourMap] = []
        self.initUI()

    def initUI(self):
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(5, 10, 5, 10)
        self.main_layout.setSpacing(10)
        self.setObjectName("IGRFWidget")

        self.mainlbl = QLabel("IGRF Column Mapping", self)
        font = self.mainlbl.font(); font.setBold(True); self.mainlbl.setFont(font)
        self.mainlbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.main_layout.addWidget(self.mainlbl)

        grid_layout = QGridLayout(); grid_layout.setVerticalSpacing(8)
        grid_layout.setHorizontalSpacing(5); grid_layout.setColumnStretch(1, 1)

        self.latlbl = QLabel("Latitude:", self)
        self.lonlbl = QLabel("Longitude:", self)
        # *** CHANGE HERE: Update label text ***
        self.altlbl = QLabel("Altitude (m):", self) # Changed from (ft) to (m)
        self.datelbl = QLabel("Date:", self)
        self.lattxt = QComboBox(self); self.lontxt = QComboBox(self)
        self.alttxt = QComboBox(self); self.datetxt = QComboBox(self)

        row = 0
        for lbl, combo in zip([self.latlbl, self.lonlbl, self.altlbl, self.datelbl],
                              [self.lattxt, self.lontxt, self.alttxt, self.datetxt]):
            lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            combo.setItemDelegate(ColorDelegate(self)); combo.setMinimumWidth(160)
            # *** CHANGE HERE: Update tooltip text to match label ***
            combo.setToolTip(f"Select the column containing {lbl.text().replace(':', '')} data")
            grid_layout.addWidget(lbl, row, 0); grid_layout.addWidget(combo, row, 1)
            row += 1
        self.main_layout.addLayout(grid_layout)

    # ... (rest of IGRFWidget methods remain the same) ...
    def refresh(self, columns: List[ColourMap]):
        """Update combo boxes, attempt guess/restore."""
        self._column_map_data = columns
        current_selections = self.get()
        self._populate_all_combo_boxes()
        self._try_restore_selections(current_selections)
        self._try_guess_columns() # Attempt to guess after populating

    def _populate_all_combo_boxes(self):
        combos = [self.lattxt, self.lontxt, self.alttxt, self.datetxt]
        for combo in combos:
            combo.blockSignals(True); combo.clear()
            model = QStandardItemModel(); item = QStandardItem(""); model.appendRow(item) # Blank default
            for col_map in self._column_map_data:
                item = QStandardItem(col_map["string"])
                item.setData(col_map["colour"], Qt.ItemDataRole.BackgroundRole)
                item.setToolTip(f"Column Name: {col_map['string']}")
                model.appendRow(item)
            combo.setModel(model); combo.setCurrentIndex(0)
            combo.blockSignals(False)

    def _try_restore_selections(self, selections: dict):
        combo_map = {"latitude": self.lattxt, "longitude": self.lontxt,
                     "altitude": self.alttxt, "date": self.datetxt}
        for key, combo in combo_map.items():
            if selections.get(key):
                index = combo.findText(selections[key])
                if index > 0: combo.setCurrentIndex(index)

    def _try_guess_columns(self):
        aliases = {
            "latitude": ["lat", "latitude", "dlat"],
            "longitude": ["lon", "long", "longitude", "dlon"],
            # *** CHANGE HERE: Update altitude aliases to favor meters ***
            "altitude": ["alt", "altitude", "elev", "elevation", "h", "height", "z", "depth", "meters", "meter", "m"], # Removed ft/feet
            "date": ["date", "time", "datetime", "timestamp", "day", "when"]
        }
        combos = {"latitude": self.lattxt, "longitude": self.lontxt,
                  "altitude": self.alttxt, "date": self.datetxt}
        assigned_cols = {combo.currentText() for combo in combos.values() if combo.currentIndex() > 0}

        for key, combo in combos.items():
            if combo.currentIndex() == 0: # Only guess if blank
                best_match_index = -1
                # Try exact match first
                for i in range(1, combo.count()):
                    col_name = combo.itemText(i)
                    if col_name.lower() == key and col_name not in assigned_cols:
                        best_match_index = i; break
                # Try aliases if no exact match
                if best_match_index == -1:
                    for alias in aliases[key]:
                        alias_lower = alias.lower()
                        for i in range(1, combo.count()):
                            col_name = combo.itemText(i); col_name_lower = col_name.lower()
                            # Simplified matching logic (exact or contains alias surrounded by non-alphanum or start/end)
                            if (alias_lower == col_name_lower or
                                f" {alias_lower} " in f" {col_name_lower} " or # Handles spaces
                                f"_{alias_lower}_" in f"_{col_name_lower}_" or # Handles underscores
                                col_name_lower.startswith(alias_lower + '_') or
                                col_name_lower.endswith('_' + alias_lower) or
                                col_name_lower.startswith(alias_lower + ' ') or # Handles space at end
                                col_name_lower.endswith(' ' + alias_lower) or # Handles space at start
                                (col_name_lower.startswith(alias_lower) and len(col_name_lower) > len(alias_lower) and not col_name_lower[len(alias_lower)].isalnum()) or # e.g. lat(deg)
                                (col_name_lower.endswith(alias_lower) and len(col_name_lower) > len(alias_lower) and not col_name_lower[-(len(alias_lower)+1)].isalnum()) # e.g. survey_lat
                               ) and col_name not in assigned_cols:
                                best_match_index = i; break
                        if best_match_index != -1: break

                if best_match_index != -1:
                    matched_text = combo.itemText(best_match_index)
                    combo.setCurrentIndex(best_match_index); assigned_cols.add(matched_text)

    def get(self) -> IGRFWidgetColumnNames:
        return IGRFWidgetColumnNames(
            latitude=self.lattxt.currentText(), longitude=self.lontxt.currentText(),
            altitude=self.alttxt.currentText(), date=self.datetxt.currentText())

    def validate_selection(self) -> bool:
        selections = self.get(); selected_values = []; missing = []
        for key, value in selections.items():
            if not value: missing.append(key.capitalize())
            else: selected_values.append(value)
        if missing:
             QMessageBox.warning(self, "Missing Selections", f"Please select columns for: {', '.join(missing)}."); return False
        if len(selected_values) != len(set(selected_values)):
             QMessageBox.warning(self, "Duplicate Selection", "Latitude, Longitude, Altitude, Date must use unique columns."); return False
        return True


# === IGRF Calculation Progress Dialog ===
class IGRFCalculator(QWidget):
    """Modal dialog showing calculation progress and handling saving."""
    dataframe: pd.DataFrame | None = None

    def __init__(self, parent=None):
        super().__init__(parent, Qt.WindowType.Dialog | Qt.WindowType.WindowStaysOnTopHint)
        self.initUI()
        self._calculation_successful = False

    def setDataFrame(self, df: pd.DataFrame | None):
        self.dataframe = df.copy() if df is not None else None

    def initUI(self):
        self.setWindowTitle("Calculating IGRF...")
        self.setMinimumWidth(450); self.setMinimumHeight(120)
        layout = QVBoxLayout(self)
        self.progressLabel = QLabel("Preparing calculation...", self)
        self.progressLabel.setAlignment(Qt.AlignmentFlag.AlignCenter); self.progressLabel.setWordWrap(True)
        self.progressBar = QProgressBar(self)
        self.progressBar.setTextVisible(True); self.progressBar.setRange(0, 100); self.progressBar.setValue(0)
        layout.addWidget(self.progressLabel); layout.addWidget(self.progressBar); layout.addStretch(1)
        # Centering deferred until just before show()

    def _center_on_parent(self):
        if self.parentWidget():
            try:
                parent_rect = self.parentWidget().geometry()
                self.adjustSize()
                self.move(parent_rect.center() - self.rect().center())
            except Exception as e: print(f"Warning: Could not center calculator dialog - {e}", file=sys.stderr)

    def run_calculation(self, columns: IGRFWidgetColumnNames) -> bool:
        """Performs calculation, updates progress, saves. Returns True if saved."""
        self._calculation_successful = False
        if self.dataframe is None or self.dataframe.empty:
            QMessageBox.critical(self, "Error", "No data available for calculation."); self.close(); return False

        num_rows = len(self.dataframe)
        if num_rows == 0:
             QMessageBox.information(self, "No Data", "Table contains no data rows."); self.close(); return False

        self.progressBar.setRange(0, num_rows); self.progressBar.setValue(0)
        self.progressLabel.setText(f"Processing row 0 of {num_rows}..."); QApplication.processEvents()

        results_igrf, results_inc, results_dec = [], [], []
        calculation_errors = 0
        first_error_details = None # Store details of the first error

        try:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            required_cols = list(columns.values())
            if not all(col in self.dataframe.columns for col in required_cols):
                missing = [col for col in required_cols if col not in self.dataframe.columns]
                raise KeyError(f"Column(s) not found: {', '.join(missing)}")

            df_subset = self.dataframe[required_cols]

            for index, row_tuple in enumerate(df_subset.itertuples(index=False)):
                row_data = dict(zip(columns.keys(), row_tuple))
                result = calculate_igrf_single(
                    row_data["latitude"], row_data["longitude"],
                    row_data["altitude"], row_data["date"]
                )

                if result and result.get("IGRF") is not None:
                    results_igrf.append(result["IGRF"]); results_inc.append(result.get("INC")); results_dec.append(result.get("DEC"))
                else:
                    results_igrf.append(pd.NA); results_inc.append(pd.NA); results_dec.append(pd.NA)
                    calculation_errors += 1
                    if first_error_details is None: # Log first error details
                        first_error_details = f"Row {index+1}: Lat={row_data.get('latitude')}, Lon={row_data.get('longitude')}, Alt={row_data.get('altitude')}, Date='{row_data.get('date')}'"

                if (index + 1) % 100 == 0 or index == num_rows - 1:
                    self.progressBar.setValue(index + 1)
                    self.progressLabel.setText(f"Processing row {index + 1} of {num_rows}...")
                    QApplication.processEvents()

        except KeyError as ke:
            QMessageBox.critical(self, "Column Error", f"Calculation failed: Column mapping error ({ke})."); self.close(); return False
        except ValueError as ve:
             index_str = f"near row {index + 1}" if 'index' in locals() else "early in data"
             QMessageBox.critical(self, "Data Error", f"Calculation stopped due to invalid data {index_str}:\n{ve}"); self.close(); return False
        except Exception as e:
            QMessageBox.critical(self, "Calc Error", f"Error during calculation ({type(e).__name__}):\n{e}"); self.close(); return False
        finally:
             QApplication.restoreOverrideCursor()

        # Add results
        try:
            igrf_col, inc_col, dec_col = 'IGRF_nT', 'INC_deg', 'DEC_deg'
            # Handle potential existing columns - overwrite is simplest
            if igrf_col in self.dataframe.columns: print(f"Warning: Overwriting existing column '{igrf_col}'")
            if inc_col in self.dataframe.columns: print(f"Warning: Overwriting existing column '{inc_col}'")
            if dec_col in self.dataframe.columns: print(f"Warning: Overwriting existing column '{dec_col}'")

            self.dataframe[igrf_col] = results_igrf
            self.dataframe[inc_col] = results_inc
            self.dataframe[dec_col] = results_dec
            self._calculation_successful = True
        except Exception as e:
             QMessageBox.critical(self, "Error Adding Columns", f"Failed add results to data:\n{e}"); self.close(); return False

        # Handle errors and save
        if calculation_errors > 0:
            error_msg = (f"{calculation_errors} / {num_rows} row(s) failed calculation.\n"
                         f"Check data validity (e.g., non-numeric values, bad dates).\n"
                         f"First error detected around:\n{first_error_details if first_error_details else '(details unavailable)'}\n\n"
                         f"These rows have NA values in the new columns.")
            QMessageBox.warning(self, "Calculation Issues", error_msg)
        elif self._calculation_successful:
             self.progressLabel.setText("Calculation complete successfully!"); QApplication.processEvents()

        self.progressLabel.setText("Choose save location..."); QApplication.processEvents()
        return self._save_results()

    def _save_results(self) -> bool:
        if not self._calculation_successful or self.dataframe is None:
            QMessageBox.critical(self, "Internal Error", "Cannot save - calculation unsuccessful."); self.close(); return False

        parent_widget = self.parentWidget(); initial_dir = os.getcwd(); suggested_filename = "calculated_data.csv"
        if parent_widget and hasattr(parent_widget, 'tableWidget') and parent_widget.tableWidget.pth:
            initial_dir = os.path.dirname(parent_widget.tableWidget.pth)
            base, ext = os.path.splitext(os.path.basename(parent_widget.tableWidget.pth))
            suggested_filename = os.path.join(initial_dir, f"{base}_IGRF{ext if ext.lower() in ['.csv', '.xls', '.xlsx'] else '.csv'}")

        save_dialog = QFileDialog(self, "Save Calculated Data", suggested_filename)
        save_dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
        save_dialog.setNameFilter("CSV files (*.csv);;Excel files (*.xlsx)") # Only offer CSV/XLSX for saving
        save_dialog.setDefaultSuffix("csv")

        if not save_dialog.exec():
            QMessageBox.information(self, "Save Cancelled", "Results were not saved."); self.close(); return False

        save_path = save_dialog.selectedFiles()[0]
        selected_filter = save_dialog.selectedNameFilter()

        try:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            save_as_excel = (".xlsx" in selected_filter or save_path.lower().endswith(".xlsx"))
            if save_as_excel:
                if not OPENPYXL_AVAILABLE:
                    QMessageBox.critical(self, "Missing Library", "Saving .xlsx requires 'openpyxl'. Install it or save as CSV."); return False
                if not save_path.lower().endswith(".xlsx"): save_path += ".xlsx"
                self.dataframe.to_excel(save_path, index=False, engine='openpyxl')
            else:
                 if not save_path.lower().endswith(".csv"): save_path += ".csv"
                 self.dataframe.to_csv(save_path, index=False, encoding='utf-8-sig') # BOM for Excel

            QMessageBox.information(self, "Save Successful", f"Data saved to:\n{save_path}")
            self.close(); return True

        except ImportError as ie: # Should be caught by OPENPYXL_AVAILABLE check, but safety
             lib_name = str(ie).split("'")[-2]
             QMessageBox.critical(self, "Missing Dependency", f"Saving requires '{lib_name}'. Install it and retry.")
        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Could not save file to {save_path}:\n({type(e).__name__}) {e}")
        finally:
            QApplication.restoreOverrideCursor()

        self.close(); return False


# === User Manual Dialog ===
class UserManualDialog(QDialog):
    """Dialog to display user manual and help information."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("IGRF Calculator - User Manual")
        self.setMinimumSize(700, 600)
        self.initUI()
    
    def initUI(self):
        layout = QVBoxLayout(self)
        
        # Create text browser for HTML content
        self.textBrowser = QTextBrowser(self)
        self.textBrowser.setOpenExternalLinks(True)
        
        # Set the manual content
        manual_html = """
        <html>
        <head>
            <style>
                body { font-family: Arial, sans-serif; margin: 20px; }
                h1 { color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }
                h2 { color: #34495e; margin-top: 20px; }
                h3 { color: #7f8c8d; }
                .section { margin-bottom: 20px; }
                .note { background-color: #fff3cd; border-left: 4px solid #ffc107; padding: 10px; margin: 10px 0; }
                .important { background-color: #f8d7da; border-left: 4px solid #dc3545; padding: 10px; margin: 10px 0; }
                code { background-color: #f4f4f4; padding: 2px 5px; border-radius: 3px; font-family: monospace; }
                ul { line-height: 1.6; }
            </style>
        </head>
        <body>
            <h1>IGRF Calculator User Manual</h1>
            
            <div class="section">
                <h2>About</h2>
                <p>The IGRF (International Geomagnetic Reference Field) Calculator computes magnetic field values
                at any location on Earth using the IGRF-14 model. The model provides main field coefficients 
                from 1900 to 2025, with secular variation extrapolations through 2030.</p>
                <p><strong>Version:</strong> IGRF-14 (2025 epoch)</p>
                <p><strong>Data Range:</strong> Years 1900-2025 with secular variation predictions to 2030</p>
            </div>
            
            <div class="section">
                <h2>Quick Start Guide</h2>
                <h3>1. Load Your Data File</h3>
                <ul>
                    <li>Click <strong>Browse...</strong> or enter the file path directly</li>
                    <li>Supported formats: CSV, XLS, XLSX</li>
                    <li>Your file should contain columns for: Latitude, Longitude, Altitude, and Date</li>
                </ul>
                
                <h3>2. Map Your Columns</h3>
                <p>After loading the file, select the appropriate columns from your data:</p>
                <ul>
                    <li><strong>Latitude:</strong> Geographic latitude in decimal degrees (-90 to 90)</li>
                    <li><strong>Longitude:</strong> Geographic longitude in decimal degrees (-180 to 180)</li>
                    <li><strong>Altitude (m):</strong> Height above sea level in meters</li>
                    <li><strong>Date:</strong> Date of the measurement (supports multiple formats)</li>
                </ul>
                
                <div class="note">
                    <strong>Note:</strong> The application will attempt to auto-detect your columns based on common naming patterns.
                </div>
                
                <h3>3. Calculate & Save</h3>
                <ul>
                    <li>Click <strong>Calculate IGRF & Save</strong> button</li>
                    <li>The calculation will run for all rows in your data</li>
                    <li>Progress will be displayed during calculation</li>
                    <li>Choose where to save the results (CSV or XLSX format)</li>
                </ul>
            </div>
            
            <div class="section">
                <h2>Output Parameters</h2>
                <p>The calculator adds three columns to your data:</p>
                <ul>
                    <li><strong>IGRF_nT:</strong> Total magnetic field intensity in nanoTesla</li>
                    <li><strong>INC_deg:</strong> Magnetic inclination (dip angle) in degrees</li>
                    <li><strong>DEC_deg:</strong> Magnetic declination in degrees</li>
                </ul>
                
                <h3>Understanding the Results</h3>
                <ul>
                    <li><strong>Total Intensity (IGRF_nT):</strong> Magnitude of the Earth's magnetic field (typical range: 20,000-70,000 nT)</li>
                    <li><strong>Declination (DEC_deg):</strong> Angle between magnetic north and true north (positive = east, negative = west)</li>
                    <li><strong>Inclination (INC_deg):</strong> Angle of the field relative to horizontal (positive = downward, negative = upward)</li>
                </ul>
            </div>
            
            <div class="section">
                <h2>Data Format Requirements</h2>
                
                <h3>Coordinate Format</h3>
                <ul>
                    <li>Latitude: Decimal degrees, -90 (South Pole) to +90 (North Pole)</li>
                    <li>Longitude: Decimal degrees, -180 (West) to +180 (East)</li>
                    <li>Altitude: Meters above sea level</li>
                </ul>
                
                <h3>Date Format</h3>
                <p>The application recognizes multiple date formats including:</p>
                <ul>
                    <li>YYYY-MM-DD (e.g., 2025-01-15)</li>
                    <li>MM/DD/YYYY (e.g., 01/15/2025)</li>
                    <li>DD/MM/YYYY (e.g., 15/01/2025)</li>
                    <li>ISO format with time (e.g., 2025-01-15T12:00:00)</li>
                </ul>
                
                <div class="important">
                    <strong>Important:</strong> Dates should be between 1900 and 2030 for best accuracy. 
                    Values outside this range will use extrapolation.
                </div>
            </div>
            
            <div class="section">
                <h2>Tips & Best Practices</h2>
                <ul>
                    <li><strong>File Size:</strong> Large files (>10,000 rows) may take several minutes to process</li>
                    <li><strong>Data Validation:</strong> Check that all rows have valid numeric coordinates and dates</li>
                    <li><strong>Missing Data:</strong> Rows with invalid data will have NA values in the output</li>
                    <li><strong>Column Names:</strong> Use descriptive column names (lat, latitude, lon, longitude, etc.) for automatic detection</li>
                    <li><strong>Save Early:</strong> For large datasets, save intermediate results to avoid data loss</li>
                </ul>
            </div>
            
            <div class="section">
                <h2>Troubleshooting</h2>
                
                <h3>File Won't Load</h3>
                <ul>
                    <li>Ensure file is in CSV, XLS, or XLSX format</li>
                    <li>Check file is not open in another program</li>
                    <li>Verify file has proper read permissions</li>
                </ul>
                
                <h3>Calculation Errors</h3>
                <ul>
                    <li>Verify latitude values are between -90 and 90</li>
                    <li>Verify longitude values are between -180 and 180</li>
                    <li>Check date format is recognized</li>
                    <li>Ensure altitude is numeric (can be 0 for sea level)</li>
                </ul>
                
                <h3>NA Values in Results</h3>
                <ul>
                    <li>Check the error count in the calculation summary</li>
                    <li>Review console output for specific error messages</li>
                    <li>Verify data types (coordinates should be numeric, not text)</li>
                </ul>
            </div>
            
            <div class="section">
                <h2>About IGRF-14</h2>
                <p>The International Geomagnetic Reference Field (IGRF) is a standard mathematical description 
                of the Earth's main magnetic field. IGRF-14 is the 14th generation model, released in 2024.</p>
                
                <p><strong>Key Features:</strong></p>
                <ul>
                    <li>Covers the period 1900-2025</li>
                    <li>Includes secular variation predictions for 2025-2030</li>
                    <li>Spherical harmonic model up to degree and order 13</li>
                    <li>Maintained by IAGA (International Association of Geomagnetism and Aeronomy)</li>
                </ul>
                
                <p><strong>References:</strong></p>
                <ul>
                    <li>Official IGRF Website: <a href="https://www.ncei.noaa.gov/products/international-geomagnetic-reference-field">NOAA/NCEI</a></li>
                    <li>IAGA Working Group: <a href="https://www.iaga-aiga.org/igrf/">IAGA V-MOD</a></li>
                </ul>
            </div>
            
            <div class="section">
                <h2>Keyboard Shortcuts</h2>
                <ul>
                    <li><strong>F1:</strong> Show this help dialog</li>
                    <li><strong>Ctrl+O:</strong> Open file dialog (when Browse button has focus)</li>
                    <li><strong>Enter:</strong> Load file (when path is entered in text box)</li>
                </ul>
            </div>
            
            <div class="section">
                <p style="text-align: center; color: #7f8c8d; margin-top: 30px;">
                    <small>IGRF Calculator | IGRF-14 (2025 epoch) | For scientific and educational use</small>
                </p>
            </div>
        </body>
        </html>
        """
        
        self.textBrowser.setHtml(manual_html)
        layout.addWidget(self.textBrowser)
        
        # Close button
        close_button = QPushButton("Close", self)
        close_button.clicked.connect(self.close)
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(close_button)
        layout.addLayout(button_layout)


# === Main Application Window ===
class Interface(QWidget):
    """Main application window using QWidget as base."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.initUI()
        # GeoMag check deferred to main() after window exists

    def initUI(self):
        self.main_layout = QGridLayout(self)
        self.setLayout(self.main_layout)

        # File Input Area (Top)
        file_input_layout = QHBoxLayout()
        self.fileNameEntry = QLineEdit(self); self.fileNameEntry.setPlaceholderText("Click Browse or enter file path and press Enter")
        self.chooseFileButton = QPushButton("Browse...", self); self.chooseFileButton.setToolTip("Select a CSV, XLS, or XLSX data file")
        file_input_layout.addWidget(self.fileNameEntry); file_input_layout.addWidget(self.chooseFileButton)
        self.main_layout.addLayout(file_input_layout, 0, 0, 1, 2) # Row 0, span 2 cols

        # Left Panel (IGRF Widget + Button)
        self.left_panel_widget = QWidget(self)
        left_panel_layout = QVBoxLayout(self.left_panel_widget)
        left_panel_layout.setContentsMargins(0, 0, 0, 0); left_panel_layout.setSpacing(15)
        self.igrfWidget = IGRFWidget(self); left_panel_layout.addWidget(self.igrfWidget)
        self.runButton = QPushButton("Calculate IGRF && Save", self); self.runButton.setToolTip("Map columns, calculate IGRF, and save results")
        runfont = self.runButton.font(); runfont.setBold(True); runfont.setPointSize(runfont.pointSize() + 1); self.runButton.setFont(runfont)
        self.runButton.setEnabled(False); self.runButton.setStyleSheet("padding: 8px 15px;")
        left_panel_layout.addWidget(self.runButton); left_panel_layout.addStretch(1)
        self.left_panel_widget.setMinimumWidth(280); self.left_panel_widget.setMaximumWidth(400)
        self.main_layout.addWidget(self.left_panel_widget, 1, 0) # Row 1, Col 0

        # Table Widget (Right Side)
        self.tableWidget = Table(self)
        self.main_layout.addWidget(self.tableWidget, 1, 1) # Row 1, Col 1

        # Grid Layout Stretching
        self.main_layout.setRowStretch(1, 1) # Middle row expands vertically
        self.main_layout.setColumnStretch(0, 0) # Left panel fixed width
        self.main_layout.setColumnStretch(1, 1) # Table expands horizontally

        # Connections
        self.chooseFileButton.clicked.connect(self.load_file_dialog)
        self.fileNameEntry.returnPressed.connect(self.load_file_from_entry)
        self.runButton.clicked.connect(self.start_igrf_calculation)

        # Initialize calculator dialog (hidden)
        self.icalc_dialog = IGRFCalculator(self)
        self.icalc_dialog.hide()

    def statusBar(self):
        """Accesses the QMainWindow's status bar."""
        parent_window = self.window()
        if isinstance(parent_window, QMainWindow) and hasattr(parent_window, 'statusBar'):
            return parent_window.statusBar()
        else: # Fallback if not in QMainWindow
            if not hasattr(self, '_dummy_status_bar'):
                class DummyStatusBar:
                    def showMessage(self, msg, timeout=0): print(f"Status: {msg}")
                self._dummy_status_bar = DummyStatusBar()
            return self._dummy_status_bar

    def load_file_dialog(self):
        if self.tableWidget.load_data(): self._update_ui_after_load()
        else: self._update_ui_after_fail()

    def load_file_from_entry(self):
        fname = self.fileNameEntry.text().strip()
        if not fname: QMessageBox.warning(self, "No Path", "Enter file path or click Browse."); return
        if self.tableWidget.load_data(fname): self._update_ui_after_load()
        else: self._update_ui_after_fail()

    def _update_ui_after_load(self):
        self.igrfWidget.refresh(self.tableWidget.color_map)
        self.fileNameEntry.setText(self.tableWidget.pth)
        self.runButton.setEnabled(self.tableWidget.df is not None and not self.tableWidget.df.empty)
        status_msg = f"Loaded: {os.path.basename(self.tableWidget.pth)}"
        if self.tableWidget.df is not None: status_msg += f" ({len(self.tableWidget.df)} rows)"
        self.statusBar().showMessage(status_msg, 5000)
        self.tableWidget.resizeColumnsToContents()

    def _update_ui_after_fail(self):
         self.runButton.setEnabled(False)
         self.igrfWidget.refresh([])
         self.statusBar().showMessage("File loading cancelled or failed.", 3000)

    def start_igrf_calculation(self):
        """Validate, show dialog, run calculation."""
        if not self.igrfWidget.validate_selection(): return
        if self.tableWidget.df is None or self.tableWidget.df.empty:
             QMessageBox.warning(self, "No Data", "Load data before calculating."); return

        selected_columns = self.igrfWidget.get()
        self.icalc_dialog.setDataFrame(self.tableWidget.df)
        self.icalc_dialog._center_on_parent() # Center before showing
        self.icalc_dialog.show()
        success = self.icalc_dialog.run_calculation(selected_columns)

        if success: self.statusBar().showMessage("Calculation complete and results saved.", 5000)
        else:
            if self.icalc_dialog._calculation_successful: # Calc done, save failed/cancelled
                 self.statusBar().showMessage("Calculation complete, results not saved.", 5000)
            else: # Calculation failed
                 self.statusBar().showMessage("Calculation failed or stopped.", 5000)

    def _check_geomag_init(self):
        """Checks GeoMag after window is created."""
        global gm
        if gm is None:
             QMessageBox.critical(self, "Initialization Error", "GeoMag library failed (check coeff.csv & console). Calculations disabled.")
             self.runButton.setEnabled(False)
             self.igrfWidget.setEnabled(False)
    
    def show_user_manual(self):
        """Display the user manual dialog."""
        manual_dialog = UserManualDialog(self)
        manual_dialog.exec()


# === Main Execution ===
def main():
    # Add script dir to path if necessary (helps PyInstaller find modules)
    script_dir_path = os.path.dirname(__file__) if "__file__" in locals() else os.getcwd()
    if script_dir_path not in sys.path: sys.path.insert(0, script_dir_path)

    app = QApplication(sys.argv)
    app.setApplicationName("IGRF Calculator")
    # app.setOrganizationName("Your Org") # Optional

    main_window = QMainWindow()
    main_window.setWindowTitle("IGRF Calculator")

    icon_path = os.path.join(script_dir_path, "icon.ico") # Or .png
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
        main_window.setWindowIcon(QIcon(icon_path))

    central_widget = Interface(main_window)
    main_window.setCentralWidget(central_widget)
    main_window.resize(1000, 750)
    main_window.statusBar() # Create status bar
    
    # Create menu bar
    menubar = main_window.menuBar()
    
    # Help menu
    help_menu = menubar.addMenu("&Help")
    
    # User Manual action
    manual_action = QAction("&User Manual", main_window)
    manual_action.setShortcut("F1")
    manual_action.setStatusTip("Show user manual and help information")
    manual_action.triggered.connect(central_widget.show_user_manual)
    help_menu.addAction(manual_action)
    
    # About action
    about_action = QAction("&About IGRF Calculator", main_window)
    about_action.setStatusTip("About this application")
    about_action.triggered.connect(lambda: QMessageBox.about(
        main_window,
        "About IGRF Calculator",
        "<h3>IGRF Calculator</h3>"
        "<p><b>Version:</b> 1.0.0 (IGRF-14)</p>"
        "<p>Calculate magnetic field values using the International Geomagnetic Reference Field model.</p>"
        "<p><b>Model:</b> IGRF-14 (2025 epoch)</p>"
        "<p><b>Data Range:</b> 1900-2025 with predictions to 2030</p>"
        "<hr>"
        "<p><small>Maintained by IAGA (International Association of Geomagnetism and Aeronomy)</small></p>"
    ))
    help_menu.addAction(about_action)
    
    main_window.show()

    # Perform GeoMag check after window exists
    central_widget._check_geomag_init()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()