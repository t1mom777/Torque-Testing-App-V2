import os
import re
import json
import threading
import pandas as pd
import serial.tools.list_ports
import openai
import tempfile
from openpyxl import load_workbook, Workbook  # For reading and generating the template

# New import for Excel to PDF conversion using win32com
try:
    import win32com.client
except ImportError:
    win32com = None

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QGridLayout, QLabel,
    QComboBox, QPushButton, QHeaderView,
    QStatusBar, QTabWidget, QTableWidget, QTableWidgetItem, QDialog,
    QFormLayout, QLineEdit, QDialogButtonBox, QHBoxLayout,
    QStackedWidget, QDoubleSpinBox, QMessageBox, QFileDialog,
    QDateEdit, QToolButton, QMenu, QApplication, QCheckBox
)
from PyQt6.QtGui import QAction, QClipboard, QImage
from PyQt6.QtCore import QThread, pyqtSignal, Qt, QDate, QTimer

from db_handler_local import (
    init_db, insert_default_torque_table_data,
    get_torque_table, insert_raw_data, insert_summary,
    add_torque_entry, update_torque_entry, delete_torque_entry,
    get_app_setting, set_app_setting
)
from serial_reader import read_from_serial, find_fits_in_selected_row
from openai_handler import perform_extraction_from_image


# ---------------- Revised Excel->PDF function ----------------
def convert_excel_to_pdf(excel_path: str, pdf_path: str):
    """
    Convert an Excel file to PDF using the Excel COM interface (pywin32).
    Opens in read-only mode, disables alerts, and ensures the workbook is closed cleanly.
    """
    if win32com is None:
        raise ImportError(
            "win32com.client module is required for Excel to PDF conversion. "
            "Please install pywin32 and run on Windows."
        )

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False  # Suppress any alerts/popups

    try:
        # Open the workbook in read-only mode to reduce locking issues
        wb = excel.Workbooks.Open(os.path.abspath(excel_path), ReadOnly=1)
        # 0 => PDF format. Use absolute path to avoid path issues.
        wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
    finally:
        # Ensure the workbook closes without saving changes
        wb.Close(SaveChanges=0)
        excel.Quit()
        # Release COM references
        del wb
        del excel


# ---------------- Worker Thread for Serial Reading ----------------
class SerialReaderWorker(QThread):
    reading_signal = pyqtSignal(float, list)

    def __init__(self, port, selected_row):
        super().__init__()
        self.port = port
        self.selected_row = selected_row
        self.stop_event = threading.Event()

    def run(self):
        BAUD_RATE = 9600

        def callback(target_torque):
            print(f"[DEBUG] Serial callback received torque: {target_torque}")
            if self.stop_event.is_set():
                return
            fits = find_fits_in_selected_row(target_torque, self.selected_row)
            if fits:
                print(f"[DEBUG] torque {target_torque} fits in ranges: {fits}")
            else:
                print(f"[DEBUG] torque {target_torque} did NOT fit any range")
            self.reading_signal.emit(target_torque, fits)

        try:
            print(f"[DEBUG] Starting serial read on {self.port} at {BAUD_RATE} baud...")
            read_from_serial(self.port, BAUD_RATE, callback, self.stop_event)
        except Exception as e:
            print("[DEBUG] Error in serial reading:", e)

    def stop(self):
        print("[DEBUG] stop_event set. Stopping serial reading.")
        self.stop_event.set()


# ---------------- Helper Functions ----------------
def calc_applied_torques(max_torque: float) -> list[float]:
    factors = [0.916, 0.583, 0.333]
    results = []
    for f in factors:
        raw_val = max_torque * f
        rounded = round(raw_val / 10) * 10
        results.append(rounded)
    return results

def calc_allowance_range(applied_val: float) -> str:
    if applied_val < 10:
        tolerance = 0.06
    else:
        tolerance = 0.04
    low = applied_val * (1 - tolerance)
    high = applied_val * (1 + tolerance)
    return f"{round(low,1)} - {round(high,1)}"


# ---------------- Helper for Filename Generation ----------------
def generate_filename(template: str, variables: dict) -> str:
    """Replace placeholders in the template (e.g. {{CustomerCompany}}) with the corresponding value."""
    filename = template
    for key, value in variables.items():
        placeholder = "{{" + key + "}}"
        filename = filename.replace(placeholder, str(value))
    return filename


# ---------------- Dialog for Adding/Editing a Torque Entry ----------------
class TorqueEntryDialog(QDialog):
    def __init__(self, parent=None, entry_data=None):
        super().__init__(parent)
        self.setWindowTitle("Torque Entry")
        self.setMinimumWidth(300)
        self.entry_data = entry_data or {}
        self.init_ui()

    def init_ui(self):
        layout = QFormLayout(self)

        self.max_torque_edit = QLineEdit(str(self.entry_data.get("max_torque", "")))
        self.unit_edit = QLineEdit(self.entry_data.get("unit", ""))
        self.type_edit = QLineEdit(self.entry_data.get("type", ""))
        self.applied_torq_edit = QLineEdit(self.entry_data.get("applied_torq", ""))
        self.allowance1_edit = QLineEdit(self.entry_data.get("allowance1", ""))
        self.allowance2_edit = QLineEdit(self.entry_data.get("allowance2", ""))
        self.allowance3_edit = QLineEdit(self.entry_data.get("allowance3", ""))

        layout.addRow("Max Torque:", self.max_torque_edit)
        layout.addRow("Unit:", self.unit_edit)
        layout.addRow("Type:", self.type_edit)
        layout.addRow("Applied Torque (JSON):", self.applied_torq_edit)
        layout.addRow("Allowance 1:", self.allowance1_edit)
        layout.addRow("Allowance 2:", self.allowance2_edit)
        layout.addRow("Allowance 3:", self.allowance3_edit)

        self.max_torque_edit.textChanged.connect(self.auto_fill_applied_from_max)
        self.applied_torq_edit.textChanged.connect(self.auto_fill_allowances_from_applied)

        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)
        self.setLayout(layout)

    def auto_fill_applied_from_max(self):
        txt = self.max_torque_edit.text().strip()
        if not txt:
            return
        match = re.search(r"[\d\.]+", txt)
        if match:
            try:
                max_torque = float(match.group())
            except ValueError:
                return
        else:
            return
        applied_list = calc_applied_torques(max_torque)
        self.applied_torq_edit.blockSignals(True)
        self.applied_torq_edit.setText(json.dumps(applied_list))
        self.applied_torq_edit.blockSignals(False)
        self.auto_fill_allowances_from_applied()

    def auto_fill_allowances_from_applied(self):
        txt = self.applied_torq_edit.text().strip()
        if not txt:
            return
        try:
            arr = json.loads(txt)
            if not isinstance(arr, list):
                return
        except (ValueError, json.JSONDecodeError):
            return
        for i in range(3):
            val = arr[i] if i < len(arr) else 0
            rng = calc_allowance_range(val)
            if i == 0:
                self.allowance1_edit.setText(rng)
            elif i == 1:
                self.allowance2_edit.setText(rng)
            else:
                self.allowance3_edit.setText(rng)

    def get_data(self):
        return {
            "max_torque": self.max_torque_edit.text().strip(),
            "unit": self.unit_edit.text().strip(),
            "type": self.type_edit.text().strip(),
            "applied_torq": self.applied_torq_edit.text().strip(),
            "allowance1": self.allowance1_edit.text().strip(),
            "allowance2": self.allowance2_edit.text().strip(),
            "allowance3": self.allowance3_edit.text().strip()
        }


# ---------------- Main Application Window ----------------
class ModernTorqueApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Torque Testing Application")
        self.setGeometry(100, 100, 950, 650)

        self.results_by_range = {}
        self.customer_info = {}
        self.serial_worker = None
        self.selected_row = None

        # OpenAI settings (defaults)
        self.openai_api_key = None
        self.openai_model = "gpt-4-turbo"
        self.openai_temperature = 0.7
        self.openai_top_p = 1.0
        self.openai_presence_penalty = 0.0
        self.openai_frequency_penalty = 0.0

        # Load saved API key and additional OpenAI settings from DB
        saved_key = get_app_setting("openai_api_key")
        if saved_key:
            self.openai_api_key = saved_key

        saved_model = get_app_setting("openai_model")
        if saved_model:
            self.openai_model = saved_model

        saved_temperature = get_app_setting("openai_temperature")
        if saved_temperature:
            try:
                self.openai_temperature = float(saved_temperature)
            except ValueError:
                pass

        saved_top_p = get_app_setting("openai_top_p")
        if saved_top_p:
            try:
                self.openai_top_p = float(saved_top_p)
            except ValueError:
                pass

        saved_presence_penalty = get_app_setting("openai_presence_penalty")
        if saved_presence_penalty:
            try:
                self.openai_presence_penalty = float(saved_presence_penalty)
            except ValueError:
                pass

        saved_frequency_penalty = get_app_setting("openai_frequency_penalty")
        if saved_frequency_penalty:
            try:
                self.openai_frequency_penalty = float(saved_frequency_penalty)
            except ValueError:
                pass

        # Load setting for showing extracted data (default hidden)
        saved_show_extracted = get_app_setting("show_extracted_data")
        if saved_show_extracted:
            self.show_extracted_data = (saved_show_extracted.lower() == "true")
        else:
            self.show_extracted_data = False

        self.setStyleSheet(self.load_stylesheet())
        self.init_ui()

    def load_stylesheet(self):
        return """
        QMainWindow {
            background-color: #FAFAFA;
            font-family: "Segoe UI", "Helvetica Neue", sans-serif;
        }
        QDialog {
            background-color: #FFFFFF;
            color: #333;
        }
        QDialog QLabel {
            color: #333;
            font-size: 14px;
        }
        QDialog QLineEdit, QDateEdit {
            background-color: #FFFFFF;
            color: #333;
            border: 1px solid #ccc;
            border-radius: 4px;
            padding: 4px;
        }
        QDialogButtonBox {
            background-color: transparent;
        }
        QTabWidget::pane {
            border: none;
            background: #FFFFFF;
        }
        QTabBar::tab {
            background: #E0E0E0;
            color: #555;
            padding: 10px 20px;
            margin: 3px;
            border-radius: 6px;
        }
        QTabBar::tab:selected {
            background: #3498db;
            color: #FFFFFF;
            font-weight: bold;
        }
        QPushButton {
            background-color: #3498db;
            border: none;
            color: #FFFFFF;
            padding: 10px 20px;
            font-size: 14px;
            border-radius: 6px;
        }
        QPushButton:disabled {
            background-color: #95a5a6;
        }
        QPushButton:hover:!pressed {
            background-color: #2980b9;
        }
        QComboBox, QLineEdit, QDateEdit {
            padding: 6px;
            font-size: 14px;
            border: 1px solid #ccc;
            border-radius: 4px;
            background-color: #FFFFFF;
            color: #333;
        }
        QComboBox QAbstractItemView {
            background-color: #F0F0F0;
            color: #333;
            selection-background-color: #3498db;
            selection-color: #FFFFFF;
        }
        QTableWidget {
            background-color: #FFFFFF;
            border: 1px solid #ccc;
            color: #333;
        }
        QHeaderView::section {
            background-color: #E0E0E0;
            color: #333;
            padding: 6px;
            border: 1px solid #ccc;
        }
        QLabel {
            font-size: 14px;
            color: #333;
        }
        """

    def init_ui(self):
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)

        self.tab_widget = QTabWidget()
        self.setCentralWidget(self.tab_widget)

        self.init_testing_tab()
        self.init_settings_tab()
        self.init_report_templates_tab()

    # ---------------- Torque Testing Tab ----------------
    def init_testing_tab(self):
        self.testing_tab = QWidget()
        main_layout = QVBoxLayout(self.testing_tab)

        # Top Info Grid
        info_grid = QGridLayout()
        row = 0

        info_grid.addWidget(QLabel("Manufacturer:"), row, 0)
        self.manufacturer_edit = QLineEdit()
        info_grid.addWidget(self.manufacturer_edit, row, 1)

        info_grid.addWidget(QLabel("Serial Number:"), row, 2)
        self.serial_number_edit = QLineEdit()
        info_grid.addWidget(self.serial_number_edit, row, 3)

        row += 1

        info_grid.addWidget(QLabel("Model:"), row, 0)
        self.model_edit = QLineEdit()
        info_grid.addWidget(self.model_edit, row, 1)

        info_grid.addWidget(QLabel("Calibration Date:"), row, 2)
        self.calibration_date_edit = QDateEdit()
        self.calibration_date_edit.setDate(QDate.currentDate())
        self.calibration_date_edit.setCalendarPopup(True)
        info_grid.addWidget(self.calibration_date_edit, row, 3)

        row += 1

        info_grid.addWidget(QLabel("Max Torque:"), row, 0)
        self.max_torque_combo = QComboBox()
        info_grid.addWidget(self.max_torque_combo, row, 1)
        self.max_torque_combo.currentIndexChanged.connect(self.on_max_torque_combo_changed)

        info_grid.addWidget(QLabel("Calibration Due:"), row, 2)
        self.calibration_due_edit = QDateEdit()
        self.calibration_due_edit.setDate(QDate.currentDate().addYears(1))
        self.calibration_due_edit.setCalendarPopup(True)
        info_grid.addWidget(self.calibration_due_edit, row, 3)

        row += 1

        info_grid.addWidget(QLabel("Unit #:"), row, 0)
        self.unit_number_edit = QLineEdit()
        info_grid.addWidget(self.unit_number_edit, row, 1)

        info_grid.addWidget(QLabel("Customer/Company:"), row, 2)
        self.customer_edit = QLineEdit()
        info_grid.addWidget(self.customer_edit, row, 3)

        row += 1

        info_grid.addWidget(QLabel("Phone Number:"), row, 0)
        self.phone_edit = QLineEdit()
        info_grid.addWidget(self.phone_edit, row, 1)

        info_grid.addWidget(QLabel("Address:"), row, 2)
        self.address_edit = QLineEdit()
        info_grid.addWidget(self.address_edit, row, 3)

        row += 1

        info_grid.addWidget(QLabel("Serial Port:"), row, 0)
        self.port_combo = QComboBox()
        self.port_combo.addItems(self.get_serial_ports())
        info_grid.addWidget(self.port_combo, row, 1)

        main_layout.addLayout(info_grid)

        # Create Test Results Table
        self.torque_table = QTableWidget()
        self.torque_table.setColumnCount(7)
        self.torque_table.setHorizontalHeaderLabels([
            "Applied Torque", "Min - Max Allowance",
            "Test 1", "Test 2", "Test 3", "Test 4", "Test 5"
        ])
        self.torque_table.setRowCount(3)
        self.torque_table.setEditTriggers(QTableWidget.EditTrigger.AllEditTriggers)
        self.torque_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        main_layout.addWidget(self.torque_table)

        self.load_max_torque_dropdown()

        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("Begin Test")
        self.start_btn.clicked.connect(self.start_test)
        btn_layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("End Test")
        self.stop_btn.clicked.connect(self.stop_test)
        self.stop_btn.setEnabled(False)
        btn_layout.addWidget(self.stop_btn)

        # Drop-down button for customer info
        self.upload_info_btn = QToolButton()
        self.upload_info_btn.setText("Import Customer Info")
        self.upload_info_btn.setPopupMode(QToolButton.ToolButtonPopupMode.MenuButtonPopup)
        menu = QMenu()
        action_upload_file = QAction("Upload image from computer", self)
        action_clipboard = QAction("Upload image/screenshot from clipboard", self)
        action_webcam = QAction("Take image from web camera", self)
        menu.addAction(action_upload_file)
        menu.addAction(action_clipboard)
        menu.addAction(action_webcam)
        self.upload_info_btn.setMenu(menu)
        action_upload_file.triggered.connect(self.upload_customer_info_from_file)
        action_clipboard.triggered.connect(self.upload_customer_info_from_clipboard)
        action_webcam.triggered.connect(self.upload_customer_info_from_webcam)
        btn_layout.addWidget(self.upload_info_btn)

        self.export_summary_btn = QPushButton("Export Summary")
        self.export_summary_btn.clicked.connect(self.export_summary)
        btn_layout.addWidget(self.export_summary_btn)

        main_layout.addLayout(btn_layout)

        self.live_torque_label = QLabel("Live Torque: --")
        self.live_torque_label.setStyleSheet("font-size: 16px; padding: 5px;")
        main_layout.addWidget(self.live_torque_label)

        # Extracted Data Section
        self.extracted_data_label = QLabel("Extracted Data:")
        main_layout.addWidget(self.extracted_data_label)
        self.extracted_data_table = QTableWidget()
        self.extracted_data_table.setColumnCount(2)
        self.extracted_data_table.setHorizontalHeaderLabels(["Field", "Value"])
        self.extracted_data_table.verticalHeader().setVisible(False)
        self.extracted_data_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        main_layout.addWidget(self.extracted_data_table)

        # Show or hide the extracted data UI based on user settings
        self.extracted_data_label.setVisible(self.show_extracted_data)
        self.extracted_data_table.setVisible(self.show_extracted_data)

        self.testing_tab.setLayout(main_layout)
        self.tab_widget.addTab(self.testing_tab, "Torque Testing")

    def get_serial_ports(self):
        ports = serial.tools.list_ports.comports()
        return [p.device for p in ports]

    def load_max_torque_dropdown(self):
        self.max_torque_combo.clear()
        table_data = get_torque_table()
        for row in table_data:
            txt = f"{row['max_torque']} {row['unit']} - {row['type']}"
            self.max_torque_combo.addItem(txt, userData=row)
        if table_data:
            self.max_torque_combo.setCurrentIndex(0)
            self.selected_row = table_data[0]
            self.display_pre_test_rows()

    def on_max_torque_combo_changed(self, index):
        row_data = self.max_torque_combo.itemData(index)
        if row_data:
            self.selected_row = row_data
            self.display_pre_test_rows()
        else:
            self.selected_row = None
            self.clear_torque_table()

    def display_pre_test_rows(self):
        self.clear_torque_table()
        if not self.selected_row:
            return
        try:
            applied_arr = json.loads(self.selected_row.get("applied_torq", "[]"))
        except json.JSONDecodeError:
            applied_arr = [0, 0, 0]
        for i in range(3):
            allowance_key = self.selected_row.get(f"allowance{i+1}", "")
            applied_val = applied_arr[i] if i < len(applied_arr) else 0
            self.torque_table.setItem(i, 0, QTableWidgetItem(str(applied_val)))
            self.torque_table.setItem(i, 1, QTableWidgetItem(allowance_key))
            for c in range(2, 7):
                self.torque_table.setItem(i, c, QTableWidgetItem(""))
        self.results_by_range = {}

    def clear_torque_table(self):
        for r in range(self.torque_table.rowCount()):
            for c in range(self.torque_table.columnCount()):
                self.torque_table.setItem(r, c, QTableWidgetItem(""))

    def start_test(self):
        if not self.selected_row:
            QMessageBox.warning(self, "Warning", "No torque entry selected.")
            return
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.statusBar.showMessage("Test in progress...")
        port = self.port_combo.currentText()
        if not port:
            QMessageBox.warning(self, "Warning", "No serial port selected.")
            return
        self.serial_worker = SerialReaderWorker(port, self.selected_row)
        self.serial_worker.reading_signal.connect(self.process_reading)
        self.serial_worker.start()

    def stop_test(self):
        if self.serial_worker:
            self.serial_worker.stop()
            self.serial_worker.wait(2000)
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.statusBar.showMessage("Test ended.")
        self.display_pre_test_rows()
        self.port_combo.clear()
        self.port_combo.addItems(self.get_serial_ports())
        QMessageBox.information(self, "Test Completed", "Test ended.")
        print("[DEBUG] Test stopped. Results by range =>", self.results_by_range)

    def process_reading(self, target_torque, fits):
        self.live_torque_label.setText(f"Live Torque: {target_torque}")
        if fits:
            self.live_torque_label.setStyleSheet("background-color: green; color: white; font-size: 16px; padding: 5px;")
        else:
            self.live_torque_label.setStyleSheet("background-color: red; color: white; font-size: 16px; padding: 5px;")
        for fit in fits:
            allowance_key = fit.get('range_str', "")
            current_results = self.results_by_range.get(allowance_key, [])
            if len(current_results) < 5:
                insert_raw_data(
                    target_torque, self.selected_row["id"],
                    f"allowance{fit.get('allowance_index', '')}",
                    allowance_key
                )
                current_results.append(target_torque)
                self.results_by_range[allowance_key] = current_results
        self.update_summary_table()

    def update_summary_table(self):
        for row_idx in range(self.torque_table.rowCount()):
            allow_item = self.torque_table.item(row_idx, 1)
            if allow_item:
                allow_key = allow_item.text().strip()
                test_vals = self.results_by_range.get(allow_key, [])
                for col_idx in range(2, 7):
                    val_index = col_idx - 2
                    if val_index < len(test_vals):
                        self.torque_table.setItem(row_idx, col_idx, QTableWidgetItem(str(test_vals[val_index])))
                    else:
                        self.torque_table.setItem(row_idx, col_idx, QTableWidgetItem(""))

    # ---------------- Export Summary (Excel + PDF) ----------------
    def export_summary(self):
        row_count = self.torque_table.rowCount()
        headers = ["Applied Torque", "Min - Max Allowance", "Test 1", "Test 2", "Test 3", "Test 4", "Test 5"]
        summary_data = []
        for r in range(row_count):
            row_dict = {}
            for c in range(self.torque_table.columnCount()):
                item = self.torque_table.item(r, c)
                row_dict[headers[c]] = item.text() if item else ""
            summary_data.append(row_dict)
        extra_info = {
            "Manufacturer": self.manufacturer_edit.text(),
            "Serial Number": self.serial_number_edit.text(),
            "Model": self.model_edit.text(),
            "Calibration Date": self.calibration_date_edit.date().toString(Qt.DateFormat.ISODate),
            "Calibration Due": self.calibration_due_edit.date().toString(Qt.DateFormat.ISODate),
            "Unit Number": self.unit_number_edit.text(),
            "Customer/Company": self.customer_edit.text(),
            "Phone Number": self.phone_edit.text(),
            "Address": self.address_edit.text(),
            "MaxTorque": f"{self.selected_row.get('max_torque', '')} {self.selected_row.get('unit', '')}" if self.selected_row else ""
        }

        if not summary_data:
            QMessageBox.warning(self, "Export Warning", "No table data to export.")
            return

        # Retrieve export settings from the DB or use defaults
        excel_save_dir = get_app_setting("excel_save_dir") or os.getcwd()
        pdf_save_dir = get_app_setting("pdf_save_dir") or os.getcwd()
        excel_filename_template = get_app_setting("excel_filename_template") or "summary_{{CustomerCompany}}_{{CalibrationDate}}.xlsx"
        pdf_filename_template = get_app_setting("pdf_filename_template") or "summary_{{CustomerCompany}}_{{CalibrationDate}}.pdf"

        # Prepare variables for filename generation
        filename_variables = {
            "Manufacturer": extra_info.get("Manufacturer", ""),
            "SerialNumber": extra_info.get("Serial Number", ""),
            "Model": extra_info.get("Model", ""),
            "CalibrationDate": extra_info.get("Calibration Date", ""),
            "CalibrationDue": extra_info.get("Calibration Due", ""),
            "UnitNumber": extra_info.get("Unit Number", ""),
            "CustomerCompany": extra_info.get("Customer/Company", ""),
            "PhoneNumber": extra_info.get("Phone Number", ""),
            "Address": extra_info.get("Address", ""),
            "MaxTorque": extra_info.get("MaxTorque", "")
        }

        excel_filename = generate_filename(excel_filename_template, filename_variables)
        pdf_filename = generate_filename(pdf_filename_template, filename_variables)
        excel_path = os.path.join(excel_save_dir, excel_filename)
        pdf_path = os.path.join(pdf_save_dir, pdf_filename)

        template_path = "summary_template.xlsx"
        try:
            if os.path.exists(template_path):
                self.export_summary_with_template(template_path, extra_info, summary_data, excel_path)
            else:
                df = pd.DataFrame(summary_data)
                df.to_excel(excel_path, index=False)

            # Convert the Excel file directly to PDF with our updated function
            convert_excel_to_pdf(excel_path, pdf_path)

            QMessageBox.information(self, "Export Summary", f"Summary exported to:\nExcel: {excel_path}\nPDF: {pdf_path}")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Error exporting summary:\n{e}")

    def export_summary_with_template(self, template_path, extra_info, summary_data, output_path):
        wb = load_workbook(template_path)
        ws = wb.active

        variables = {
            "Manufacturer": extra_info.get("Manufacturer", ""),
            "SerialNumber": extra_info.get("Serial Number", ""),
            "Model": extra_info.get("Model", ""),
            "CalibrationDate": extra_info.get("Calibration Date", ""),
            "CalibrationDue": extra_info.get("Calibration Due", ""),
            "UnitNumber": extra_info.get("Unit Number", ""),
            "CustomerCompany": extra_info.get("Customer/Company", ""),
            "PhoneNumber": extra_info.get("Phone Number", ""),
            "Address": extra_info.get("Address", ""),
            "MaxTorque": extra_info.get("MaxTorque", "")
        }

        # Replace placeholders in the template
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    for key, val in variables.items():
                        placeholder = "{{" + key + "}}"
                        if placeholder in cell.value:
                            cell.value = cell.value.replace(placeholder, str(val))

        summary_variables = {}
        for idx, row_data in enumerate(summary_data):
            allowance_number = idx + 1
            summary_variables[f"AppliedTorque{allowance_number}"] = row_data.get("Applied Torque", "")
            summary_variables[f"MinMaxAllowance{allowance_number}"] = row_data.get("Min - Max Allowance", "")
            for test in range(1, 6):
                summary_variables[f"Test{test}_Allowance{allowance_number}"] = row_data.get(f"Test {test}", "")

        # Replace placeholders for each test row
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    for key, val in summary_variables.items():
                        placeholder = "{{" + key + "}}"
                        if placeholder in cell.value:
                            cell.value = cell.value.replace(placeholder, str(val))

        wb.save(output_path)

    # ---------------- Extraction from Image via ChatGPT API ----------------
    def upload_customer_info_from_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Image", "",
            "Image Files (*.png *.jpg *.jpeg *.bmp);;All Files (*)"
        )
        if not file_path:
            return
        if not self.openai_api_key:
            QMessageBox.critical(self, "Error", "OpenAI API Key not set.")
            return
        extracted_data = self.extract_torque_data(file_path)
        if not extracted_data:
            QMessageBox.warning(self, "Extraction Failed", "No data extracted or an error occurred.")
            return
        self.update_extracted_data_table(extracted_data)

    def upload_customer_info_from_clipboard(self):
        clipboard = QApplication.clipboard()
        image = clipboard.image()
        if image.isNull():
            QMessageBox.warning(self, "Clipboard Empty", "No image found in clipboard.")
            return
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        temp_file.close()
        if not image.save(temp_file.name, "PNG"):
            QMessageBox.critical(self, "Error", "Failed to save clipboard image.")
            return
        if not self.openai_api_key:
            QMessageBox.critical(self, "Error", "OpenAI API Key not set.")
            return
        extracted_data = self.extract_torque_data(temp_file.name)
        if not extracted_data:
            QMessageBox.warning(self, "Extraction Failed", "No data extracted or an error occurred.")
            return
        self.update_extracted_data_table(extracted_data)

    def upload_customer_info_from_webcam(self):
        try:
            import cv2
        except ImportError:
            QMessageBox.critical(self, "Error", "OpenCV is not installed. Please install opencv-python.")
            return
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            QMessageBox.critical(self, "Error", "Could not open web camera.")
            return
        cv2.namedWindow("Webcam - Press Space to Capture", cv2.WINDOW_NORMAL)
        while True:
            ret, frame = cap.read()
            if not ret:
                QMessageBox.critical(self, "Error", "Failed to capture image from web camera.")
                cap.release()
                cv2.destroyAllWindows()
                return
            cv2.imshow("Webcam - Press Space to Capture", frame)
            key = cv2.waitKey(1) & 0xFF
            if key == 32:  # space key
                break
            elif key == 27:  # escape key to cancel
                cap.release()
                cv2.destroyAllWindows()
                return
        cap.release()
        cv2.destroyAllWindows()
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        temp_file.close()
        cv2.imwrite(temp_file.name, frame)
        if not self.openai_api_key:
            QMessageBox.critical(self, "Error", "OpenAI API Key not set.")
            return
        extracted_data = self.extract_torque_data(temp_file.name)
        if not extracted_data:
            QMessageBox.warning(self, "Extraction Failed", "No data extracted or an error occurred.")
            return
        self.update_extracted_data_table(extracted_data)

    def extract_torque_data(self, image_path: str) -> dict:
        return perform_extraction_from_image(image_path, self.openai_api_key, self.openai_model)

    def update_extracted_data_table(self, data: dict):
        self.manufacturer_edit.setText(data.get("manufacturer", ""))
        self.model_edit.setText(data.get("model", ""))
        self.unit_number_edit.setText(data.get("unit", ""))
        self.serial_number_edit.setText(data.get("serial", ""))
        self.customer_edit.setText(data.get("customer", ""))
        self.phone_edit.setText(data.get("phone", ""))
        self.address_edit.setText(data.get("address", ""))

        max_torque_raw = data.get("max_torque", "")
        torque_unit_str = str(data.get("torque_unit", "")).strip()

        max_torque_str = str(max_torque_raw).strip()
        if max_torque_str:
            try:
                extracted_val = float(max_torque_str)
                self.auto_select_max_torque(extracted_val, torque_unit_str)
            except ValueError:
                pass

        fields = [
            ("Manufacturer", data.get("manufacturer", "")),
            ("Model", data.get("model", "")),
            ("Unit #", data.get("unit", "")),
            ("Serial Number", data.get("serial", "")),
            ("Customer/Company", data.get("customer", "")),
            ("Phone Number", data.get("phone", "")),
            ("Address", data.get("address", "")),
            ("Max Torque", max_torque_str),
            ("Torque Unit", torque_unit_str)
        ]
        self.extracted_data_table.setRowCount(len(fields))
        for i, (field, value) in enumerate(fields):
            self.extracted_data_table.setItem(i, 0, QTableWidgetItem(field))
            self.extracted_data_table.setItem(i, 1, QTableWidgetItem(value))

    def auto_select_max_torque(self, extracted_val: float, extracted_unit: str):
        FT_LB_SYNONYMS = {
            "ft/lb", "ft-lb", "ft.lb", "ft lb",
            "ft/lbs", "ft-lbs", "ft.lbs", "ft lbs"
        }
        IN_LB_SYNONYMS = {
            "in/lb", "in-lb", "in.lb", "in lb",
            "in/lbs", "in-lbs", "in.lbs", "in lbs"
        }
        NM_SYNONYMS = {
            "nm", "n.m", "n*m", "nm.", "n.m."
        }

        def ftlb_to_nm(val):
            return val * 1.35582

        def inlb_to_nm(val):
            return val * 0.113

        extracted_unit_lower = extracted_unit.lower().strip()

        if extracted_unit_lower in FT_LB_SYNONYMS:
            extracted_val_nm = ftlb_to_nm(extracted_val)
        elif extracted_unit_lower in IN_LB_SYNONYMS:
            extracted_val_nm = inlb_to_nm(extracted_val)
        elif extracted_unit_lower in NM_SYNONYMS:
            extracted_val_nm = extracted_val
        else:
            extracted_val_nm = extracted_val

        table_data = get_torque_table()
        tolerance_base = max(extracted_val_nm * 0.05, 1.0)

        for i, row in enumerate(table_data):
            db_torque = row["max_torque"]
            db_unit_lower = row["unit"].lower().strip()
            if db_unit_lower in FT_LB_SYNONYMS:
                db_torque_nm = ftlb_to_nm(db_torque)
            elif db_unit_lower in IN_LB_SYNONYMS:
                db_torque_nm = inlb_to_nm(db_torque)
            elif db_unit_lower in NM_SYNONYMS:
                db_torque_nm = db_torque
            else:
                db_torque_nm = db_torque

            if abs(db_torque_nm - extracted_val_nm) <= tolerance_base:
                self.max_torque_combo.setCurrentIndex(i)
                self.selected_row = row
                self.display_pre_test_rows()
                return

    # ---------------- Settings Tab ----------------
    def init_settings_tab(self):
        self.settings_tab = QWidget()
        layout = QVBoxLayout(self.settings_tab)

        self.settings_combo = QComboBox()
        self.settings_combo.addItem("Data Management")
        self.settings_combo.addItem("OpenAI Settings")
        self.settings_combo.addItem("Export Settings")
        self.settings_combo.currentIndexChanged.connect(self.on_settings_combo_changed)
        layout.addWidget(self.settings_combo)

        self.settings_stacked = QStackedWidget()
        layout.addWidget(self.settings_stacked)

        # Data Management Page
        self.data_management_page = QWidget()
        dm_layout = QVBoxLayout(self.data_management_page)

        self.torque_table_widget = QTableWidget()
        self.torque_table_widget.setColumnCount(4)
        self.torque_table_widget.setHorizontalHeaderLabels([
            "Max Torque", "Unit", "Type", "Applied Torque"
        ])
        self.torque_table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        dm_layout.addWidget(self.torque_table_widget)

        btn_layout = QHBoxLayout()
        self.add_btn = QPushButton("Add Entry")
        self.add_btn.clicked.connect(self.add_entry)
        self.edit_btn = QPushButton("Edit Entry")
        self.edit_btn.clicked.connect(self.edit_entry)
        self.delete_btn = QPushButton("Delete Entry")
        self.delete_btn.clicked.connect(self.delete_entry)
        self.refresh_btn = QPushButton("Refresh")
        self.refresh_btn.clicked.connect(self.load_torque_table_data)
        btn_layout.addWidget(self.add_btn)
        btn_layout.addWidget(self.edit_btn)
        btn_layout.addWidget(self.delete_btn)
        btn_layout.addWidget(self.refresh_btn)
        dm_layout.addLayout(btn_layout)

        self.extracted_data_checkbox = QCheckBox("Show Extracted Data in Testing Tab")
        self.extracted_data_checkbox.setChecked(self.show_extracted_data)
        self.extracted_data_checkbox.stateChanged.connect(self.toggle_extracted_data)
        dm_layout.addWidget(self.extracted_data_checkbox)

        self.data_management_page.setLayout(dm_layout)
        self.settings_stacked.addWidget(self.data_management_page)

        # OpenAI Settings Page
        self.openai_settings_page = QWidget()
        openai_layout = QFormLayout(self.openai_settings_page)

        self.api_key_edit = QLineEdit()
        if self.openai_api_key:
            self.api_key_edit.setText(self.openai_api_key)
        openai_layout.addRow("OpenAI API Key:", self.api_key_edit)

        self.model_combo = QComboBox()
        self.model_combo.addItems([
            "gpt-4o",
            "gpt-4o-mini",
            "gpt-4-turbo"
        ])
        self.model_combo.setCurrentText(self.openai_model)
        openai_layout.addRow("Model:", self.model_combo)

        self.temp_spin = QDoubleSpinBox()
        self.temp_spin.setRange(0.0, 2.0)
        self.temp_spin.setSingleStep(0.1)
        self.temp_spin.setValue(self.openai_temperature)
        openai_layout.addRow("Temperature:", self.temp_spin)

        self.top_p_spin = QDoubleSpinBox()
        self.top_p_spin.setRange(0.0, 1.0)
        self.top_p_spin.setSingleStep(0.1)
        self.top_p_spin.setValue(self.openai_top_p)
        openai_layout.addRow("Top P:", self.top_p_spin)

        self.presence_spin = QDoubleSpinBox()
        self.presence_spin.setRange(0.0, 2.0)
        self.presence_spin.setSingleStep(0.1)
        self.presence_spin.setValue(self.openai_presence_penalty)
        openai_layout.addRow("Presence Penalty:", self.presence_spin)

        self.freq_spin = QDoubleSpinBox()
        self.freq_spin.setRange(0.0, 2.0)
        self.freq_spin.setSingleStep(0.1)
        self.freq_spin.setValue(self.openai_frequency_penalty)
        openai_layout.addRow("Frequency Penalty:", self.freq_spin)

        save_key_btn = QPushButton("Save OpenAI Settings")
        save_key_btn.clicked.connect(self.save_openai_settings)
        openai_layout.addWidget(save_key_btn)

        self.settings_stacked.addWidget(self.openai_settings_page)

        # Export Settings Page
        self.export_settings_page = QWidget()
        export_layout = QFormLayout(self.export_settings_page)

        self.filename_vars = [
            ("Manufacturer", "{{Manufacturer}}"),
            ("Serial Number", "{{SerialNumber}}"),
            ("Model", "{{Model}}"),
            ("Calibration Date", "{{CalibrationDate}}"),
            ("Max Torque", "{{MaxTorque}}"),
            ("Calibration Due", "{{CalibrationDue}}"),
            ("Unit Number", "{{UnitNumber}}"),
            ("Customer/Company", "{{CustomerCompany}}"),
            ("Phone Number", "{{PhoneNumber}}"),
            ("Address", "{{Address}}"),
        ]

        self.excel_dir_edit = QLineEdit()
        self.excel_dir_edit.setText(get_app_setting("excel_save_dir") or os.getcwd())
        excel_browse_btn = QPushButton("Browse")
        excel_browse_btn.clicked.connect(self.browse_excel_dir)
        hbox_excel = QHBoxLayout()
        hbox_excel.addWidget(self.excel_dir_edit)
        hbox_excel.addWidget(excel_browse_btn)
        export_layout.addRow("Excel Save Directory:", hbox_excel)

        self.pdf_dir_edit = QLineEdit()
        self.pdf_dir_edit.setText(get_app_setting("pdf_save_dir") or os.getcwd())
        pdf_browse_btn = QPushButton("Browse")
        pdf_browse_btn.clicked.connect(self.browse_pdf_dir)
        hbox_pdf = QHBoxLayout()
        hbox_pdf.addWidget(self.pdf_dir_edit)
        hbox_pdf.addWidget(pdf_browse_btn)
        export_layout.addRow("PDF Save Directory:", hbox_pdf)

        # Excel filename template + variable combo
        excel_template_layout = QHBoxLayout()
        self.excel_template_edit = QLineEdit()
        self.excel_template_edit.setText(
            get_app_setting("excel_filename_template") or "summary_{{CustomerCompany}}_{{CalibrationDate}}.xlsx"
        )
        excel_template_layout.addWidget(self.excel_template_edit)

        self.excel_var_combo = QComboBox()
        self.excel_var_combo.addItem("-- Insert variable --")
        for label, placeholder in self.filename_vars:
            self.excel_var_combo.addItem(label, placeholder)
        self.excel_var_combo.currentIndexChanged.connect(self.on_excel_var_changed)
        excel_template_layout.addWidget(self.excel_var_combo)
        export_layout.addRow("Excel Filename Template:", excel_template_layout)

        # PDF filename template + variable combo
        pdf_template_layout = QHBoxLayout()
        self.pdf_template_edit = QLineEdit()
        self.pdf_template_edit.setText(
            get_app_setting("pdf_filename_template") or "summary_{{CustomerCompany}}_{{CalibrationDate}}.pdf"
        )
        pdf_template_layout.addWidget(self.pdf_template_edit)

        self.pdf_var_combo = QComboBox()
        self.pdf_var_combo.addItem("-- Insert variable --")
        for label, placeholder in self.filename_vars:
            self.pdf_var_combo.addItem(label, placeholder)
        self.pdf_var_combo.currentIndexChanged.connect(self.on_pdf_var_changed)
        pdf_template_layout.addWidget(self.pdf_var_combo)
        export_layout.addRow("PDF Filename Template:", pdf_template_layout)

        save_export_btn = QPushButton("Save Export Settings")
        save_export_btn.clicked.connect(self.save_export_settings)
        export_layout.addWidget(save_export_btn)

        self.export_settings_page.setLayout(export_layout)
        self.settings_stacked.addWidget(self.export_settings_page)

        self.settings_tab.setLayout(layout)
        self.tab_widget.addTab(self.settings_tab, "Settings")

        self.load_torque_table_data()

    def on_settings_combo_changed(self, index):
        self.settings_stacked.setCurrentIndex(index)

    def on_excel_var_changed(self, index):
        if index <= 0:
            return
        var_str = self.excel_var_combo.itemData(index)
        current_text = self.excel_template_edit.text()
        self.excel_template_edit.setText(current_text + var_str)
        self.excel_var_combo.setCurrentIndex(0)

    def on_pdf_var_changed(self, index):
        if index <= 0:
            return
        var_str = self.pdf_var_combo.itemData(index)
        current_text = self.pdf_template_edit.text()
        self.pdf_template_edit.setText(current_text + var_str)
        self.pdf_var_combo.setCurrentIndex(0)

    def browse_excel_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Excel Save Directory")
        if directory:
            self.excel_dir_edit.setText(directory)

    def browse_pdf_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "Select PDF Save Directory")
        if directory:
            self.pdf_dir_edit.setText(directory)

    def save_export_settings(self):
        set_app_setting("excel_save_dir", self.excel_dir_edit.text().strip())
        set_app_setting("pdf_save_dir", self.pdf_dir_edit.text().strip())
        set_app_setting("excel_filename_template", self.excel_template_edit.text().strip())
        set_app_setting("pdf_filename_template", self.pdf_template_edit.text().strip())
        QMessageBox.information(self, "Export Settings", "Export settings saved.")

    def toggle_extracted_data(self, state):
        visible = (state == Qt.CheckState.Checked)
        self.show_extracted_data = visible
        set_app_setting("show_extracted_data", "true" if visible else "false")
        self.extracted_data_label.setVisible(visible)
        self.extracted_data_table.setVisible(visible)

    def save_openai_settings(self):
        self.openai_api_key = self.api_key_edit.text().strip()
        self.openai_model = self.model_combo.currentText()
        self.openai_temperature = self.temp_spin.value()
        self.openai_top_p = self.top_p_spin.value()
        self.openai_presence_penalty = self.presence_spin.value()
        self.openai_frequency_penalty = self.freq_spin.value()
        set_app_setting("openai_api_key", self.openai_api_key)
        set_app_setting("openai_model", self.openai_model)
        set_app_setting("openai_temperature", str(self.openai_temperature))
        set_app_setting("openai_top_p", str(self.openai_top_p))
        set_app_setting("openai_presence_penalty", str(self.openai_presence_penalty))
        set_app_setting("openai_frequency_penalty", str(self.openai_frequency_penalty))
        QMessageBox.information(self, "OpenAI Settings", "OpenAI settings saved in DB.")

    def load_torque_table_data(self):
        table_data = get_torque_table()
        self.torque_table_widget.setRowCount(len(table_data))
        for i, row in enumerate(table_data):
            self.torque_table_widget.setItem(i, 0, QTableWidgetItem(str(row.get("max_torque", ""))))
            self.torque_table_widget.setItem(i, 1, QTableWidgetItem(row.get("unit", "")))
            self.torque_table_widget.setItem(i, 2, QTableWidgetItem(row.get("type", "")))
            self.torque_table_widget.setItem(i, 3, QTableWidgetItem(row.get("applied_torq", "")))
        self.torque_table_widget.resizeColumnsToContents()

    def add_entry(self):
        dialog = TorqueEntryDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.get_data()
            add_torque_entry(
                data["max_torque"], data["unit"], data["type"],
                data["applied_torq"], data["allowance1"],
                data["allowance2"], data["allowance3"]
            )
            QMessageBox.information(self, "Success", "Entry added successfully.")
            self.load_torque_table_data()
            self.load_max_torque_dropdown()

    def edit_entry(self):
        selected_items = self.torque_table_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select an entry to edit.")
            return
        row = self.torque_table_widget.currentRow()
        table_data = get_torque_table()
        entry = table_data[row]
        dialog = TorqueEntryDialog(self, entry)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.get_data()
            update_torque_entry(
                entry["id"], data["max_torque"], data["unit"], data["type"],
                data["applied_torq"], data["allowance1"], data["allowance2"], data["allowance3"]
            )
            QMessageBox.information(self, "Success", "Entry updated successfully.")
            self.load_torque_table_data()
            self.load_max_torque_dropdown()

    def delete_entry(self):
        selected_items = self.torque_table_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select an entry to delete.")
            return
        row = self.torque_table_widget.currentRow()
        table_data = get_torque_table()
        entry = table_data[row]
        confirm = QMessageBox.question(
            self, "Confirm Deletion", f"Are you sure you want to delete entry ID {entry['id']}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if confirm == QMessageBox.StandardButton.Yes:
            delete_torque_entry(entry["id"])
            QMessageBox.information(self, "Success", "Entry deleted successfully.")
            self.load_torque_table_data()
            self.load_max_torque_dropdown()

    def init_report_templates_tab(self):
        self.report_templates_tab = QWidget()
        layout = QVBoxLayout(self.report_templates_tab)
        self.template_editor_btn = QPushButton("Open Report Template Editor")
        self.template_editor_btn.clicked.connect(self.open_template_editor)
        layout.addWidget(self.template_editor_btn)
        self.generate_template_btn = QPushButton("Generate Summary Template")
        self.generate_template_btn.clicked.connect(self.generate_summary_template)
        layout.addWidget(self.generate_template_btn)
        self.report_templates_tab.setLayout(layout)
        self.tab_widget.addTab(self.report_templates_tab, "Report Templates")

    def open_template_editor(self):
        from template_editor import TemplateEditor
        self.template_editor = TemplateEditor()
        self.template_editor.show()

    def generate_summary_template(self):
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Summary Template"
            ws.cell(row=1, column=1, value="Manufacturer:")
            ws.cell(row=1, column=2, value="{{Manufacturer}}")
            ws.cell(row=2, column=1, value="Serial Number:")
            ws.cell(row=2, column=2, value="{{SerialNumber}}")
            ws.cell(row=3, column=1, value="Model:")
            ws.cell(row=3, column=2, value="{{Model}}")
            ws.cell(row=4, column=1, value="Calibration Date:")
            ws.cell(row=4, column=2, value="{{CalibrationDate}}")
            ws.cell(row=5, column=1, value="Calibration Due:")
            ws.cell(row=5, column=2, value="{{CalibrationDue}}")
            ws.cell(row=6, column=1, value="Unit Number:")
            ws.cell(row=6, column=2, value="{{UnitNumber}}")
            ws.cell(row=7, column=1, value="Customer/Company:")
            ws.cell(row=7, column=2, value="{{CustomerCompany}}")
            ws.cell(row=8, column=1, value="Phone Number:")
            ws.cell(row=8, column=2, value="{{PhoneNumber}}")
            ws.cell(row=9, column=1, value="Address:")
            ws.cell(row=9, column=2, value="{{Address}}")
            ws.cell(row=10, column=1, value="Max Torque:")
            ws.cell(row=10, column=2, value="{{MaxTorque}}")

            start_table = 12
            headers = ["Allowance", "Applied Torque", "Min - Max Allowance", "Test 1", "Test 2", "Test 3", "Test 4", "Test 5"]
            for col, header in enumerate(headers, start=1):
                ws.cell(row=start_table, column=col, value=header)

            wb.save("summary_template.xlsx")
            QMessageBox.information(self, "Template Generated", "Summary template generated as 'summary_template.xlsx'.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate template:\n{e}")

    def extract_torque_data(self, image_path: str) -> dict:
        return perform_extraction_from_image(image_path, self.openai_api_key, self.openai_model)

    def update_extracted_data_table(self, data: dict):
        self.manufacturer_edit.setText(data.get("manufacturer", ""))
        self.model_edit.setText(data.get("model", ""))
        self.unit_number_edit.setText(data.get("unit", ""))
        self.serial_number_edit.setText(data.get("serial", ""))
        self.customer_edit.setText(data.get("customer", ""))
        self.phone_edit.setText(data.get("phone", ""))
        self.address_edit.setText(data.get("address", ""))

        max_torque_raw = data.get("max_torque", "")
        torque_unit_str = str(data.get("torque_unit", "")).strip()

        max_torque_str = str(max_torque_raw).strip()
        if max_torque_str:
            try:
                extracted_val = float(max_torque_str)
                self.auto_select_max_torque(extracted_val, torque_unit_str)
            except ValueError:
                pass

        fields = [
            ("Manufacturer", data.get("manufacturer", "")),
            ("Model", data.get("model", "")),
            ("Unit #", data.get("unit", "")),
            ("Serial Number", data.get("serial", "")),
            ("Customer/Company", data.get("customer", "")),
            ("Phone Number", data.get("phone", "")),
            ("Address", data.get("address", "")),
            ("Max Torque", max_torque_str),
            ("Torque Unit", torque_unit_str)
        ]
        self.extracted_data_table.setRowCount(len(fields))
        for i, (field, value) in enumerate(fields):
            self.extracted_data_table.setItem(i, 0, QTableWidgetItem(field))
            self.extracted_data_table.setItem(i, 1, QTableWidgetItem(value))

    def auto_select_max_torque(self, extracted_val: float, extracted_unit: str):
        FT_LB_SYNONYMS = {
            "ft/lb", "ft-lb", "ft.lb", "ft lb",
            "ft/lbs", "ft-lbs", "ft.lbs", "ft lbs"
        }
        IN_LB_SYNONYMS = {
            "in/lb", "in-lb", "in.lb", "in lb",
            "in/lbs", "in-lbs", "in.lbs", "in lbs"
        }
        NM_SYNONYMS = {
            "nm", "n.m", "n*m", "nm.", "n.m."
        }

        def ftlb_to_nm(val):
            return val * 1.35582

        def inlb_to_nm(val):
            return val * 0.113

        extracted_unit_lower = extracted_unit.lower().strip()

        if extracted_unit_lower in FT_LB_SYNONYMS:
            extracted_val_nm = ftlb_to_nm(extracted_val)
        elif extracted_unit_lower in IN_LB_SYNONYMS:
            extracted_val_nm = inlb_to_nm(extracted_val)
        elif extracted_unit_lower in NM_SYNONYMS:
            extracted_val_nm = extracted_val
        else:
            extracted_val_nm = extracted_val

        table_data = get_torque_table()
        tolerance_base = max(extracted_val_nm * 0.05, 1.0)

        for i, row in enumerate(table_data):
            db_torque = row["max_torque"]
            db_unit_lower = row["unit"].lower().strip()
            if db_unit_lower in FT_LB_SYNONYMS:
                db_torque_nm = ftlb_to_nm(db_torque)
            elif db_unit_lower in IN_LB_SYNONYMS:
                db_torque_nm = inlb_to_nm(db_torque)
            elif db_unit_lower in NM_SYNONYMS:
                db_torque_nm = db_torque
            else:
                db_torque_nm = db_torque

            if abs(db_torque_nm - extracted_val_nm) <= tolerance_base:
                self.max_torque_combo.setCurrentIndex(i)
                self.selected_row = row
                self.display_pre_test_rows()
                return

def main():
    import sys
    app = QApplication(sys.argv)
    init_db()
    insert_default_torque_table_data()
    window = ModernTorqueApp()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
