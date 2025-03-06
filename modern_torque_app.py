import os
import re
import json
import threading
import pandas as pd
import serial.tools.list_ports
import openai
import tempfile

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QGridLayout, QLabel,
    QComboBox, QPushButton, QTableWidget, QHeaderView,
    QStatusBar, QTabWidget, QTableWidgetItem, QDialog,
    QFormLayout, QLineEdit, QDialogButtonBox, QHBoxLayout,
    QStackedWidget, QDoubleSpinBox, QMessageBox, QFileDialog,
    QDateEdit, QToolButton, QMenu, QApplication, QCheckBox
)
from PyQt6.QtGui import QAction, QClipboard, QImage
from PyQt6.QtCore import QThread, pyqtSignal, Qt, QDate, QTimer

# Import DB handler (which now includes AppSettings functions)
from db_handler_local import (
    get_torque_table, insert_raw_data, insert_summary,
    add_torque_entry, update_torque_entry, delete_torque_entry,
    get_app_setting, set_app_setting
)
from serial_reader import read_from_serial, find_fits_in_selected_row

# Import the extraction function from openai_handler.
from openai_handler import perform_extraction_from_image


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

        self.export_excel_btn = QPushButton("Export Summary")
        self.export_excel_btn.clicked.connect(self.export_summary_to_excel)
        btn_layout.addWidget(self.export_excel_btn)

        main_layout.addLayout(btn_layout)

        self.live_torque_label = QLabel("Live Torque: --")
        self.live_torque_label.setStyleSheet("font-size: 16px; padding: 5px;")
        main_layout.addWidget(self.live_torque_label)

        # ---------------- Extracted Data Section ----------------
        self.extracted_data_label = QLabel("Extracted Data:")
        main_layout.addWidget(self.extracted_data_label)
        self.extracted_data_table = QTableWidget()
        self.extracted_data_table.setColumnCount(2)
        self.extracted_data_table.setHorizontalHeaderLabels(["Field", "Value"])
        self.extracted_data_table.verticalHeader().setVisible(False)
        self.extracted_data_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        main_layout.addWidget(self.extracted_data_table)
        # Set visibility based on the setting
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

    def export_summary_to_excel(self):
        row_count = self.torque_table.rowCount()
        headers = ["Applied Torque", "Min - Max Allowance", "Test 1", "Test 2", "Test 3", "Test 4", "Test 5"]
        summary_data = []
        for r in range(row_count):
            row_dict = {}
            for c in range(self.torque_table.columnCount()):
                item = self.torque_table.item(r, c)
                row_dict[headers[c]] = item.text() if item else ""
            summary_data.append(row_dict)
        # Additional customer/test info
        extra_info = {
            "Manufacturer": self.manufacturer_edit.text(),
            "Serial Number": self.serial_number_edit.text(),
            "Model": self.model_edit.text(),
            "Calibration Date": self.calibration_date_edit.date().toString(Qt.DateFormat.ISODate),
            "Calibration Due": self.calibration_due_edit.date().toString(Qt.DateFormat.ISODate),
            "Unit Number": self.unit_number_edit.text(),
            "Customer/Company": self.customer_edit.text(),
            "Phone Number": self.phone_edit.text(),
            "Address": self.address_edit.text()
        }
        for row in summary_data:
            row.update(extra_info)
        if summary_data:
            df = pd.DataFrame(summary_data)
            try:
                df.to_excel("summary.xlsx", index=False)
                QMessageBox.information(self, "Export Summary", "Summary exported to summary.xlsx")
            except Exception as e:
                QMessageBox.critical(self, "Export Error", f"Error exporting to Excel:\n{e}")
        else:
            QMessageBox.warning(self, "Export Warning", "No table data to export.")

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
        # Save image to a temporary file
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
        # Save the captured image to a temporary file
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
        """
        Uses the ChatGPT API to extract specific torque wrench details from the image.
        Expects a JSON response with keys:
        manufacturer, model, unit, serial, customer, phone, address, max_torque, torque_unit
        """
        return perform_extraction_from_image(image_path, self.openai_api_key, self.openai_model)

    def update_extracted_data_table(self, data: dict):
        # Update the top info fields with extracted data
        self.manufacturer_edit.setText(data.get("manufacturer", ""))
        self.model_edit.setText(data.get("model", ""))
        self.unit_number_edit.setText(data.get("unit", ""))
        self.serial_number_edit.setText(data.get("serial", ""))
        self.customer_edit.setText(data.get("customer", ""))
        self.phone_edit.setText(data.get("phone", ""))
        self.address_edit.setText(data.get("address", ""))

        # Safely handle max_torque (which might be int/float/string)
        max_torque_raw = data.get("max_torque", "")
        torque_unit_str = str(data.get("torque_unit", "")).strip()

        # Convert to string before strip
        max_torque_str = str(max_torque_raw).strip()
        if max_torque_str:
            try:
                extracted_val = float(max_torque_str)
                self.auto_select_max_torque(extracted_val, torque_unit_str)
            except ValueError:
                pass

        # Update the bottom extracted data table
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

    #
    # -------------- EXPANDED auto_select_max_torque --------------
    #
    def auto_select_max_torque(self, extracted_val: float, extracted_unit: str):
        """
        Tries to match the extracted torque and unit with an existing row in the database.
        Supports multiple synonyms for ft-lb, in-lb, and Nm (e.g. 'ft lb', 'ft-lbs', etc.).
        If found, auto-select it in the combo box. Otherwise, user can still select manually.
        """

        # Define sets of recognized synonyms for each unit
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

        # Helper conversions to Nm
        def ftlb_to_nm(val):
            return val * 1.35582

        def inlb_to_nm(val):
            return val * 0.113  # 1 in-lb ~ 0.113 Nm

        # Convert the extracted unit to a canonical form (just for matching)
        extracted_unit_lower = extracted_unit.lower().strip()

        # 1) Determine the "extracted_val" in Nm
        if extracted_unit_lower in FT_LB_SYNONYMS:
            extracted_val_nm = ftlb_to_nm(extracted_val)
        elif extracted_unit_lower in IN_LB_SYNONYMS:
            extracted_val_nm = inlb_to_nm(extracted_val)
        elif extracted_unit_lower in NM_SYNONYMS:
            extracted_val_nm = extracted_val  # already in Nm
        else:
            # If it's an unknown unit, we can't convert => just do a direct float match
            extracted_val_nm = extracted_val

        # Retrieve all rows
        table_data = get_torque_table()

        # We'll match if the difference is within a tolerance (5% or at least 1.0)
        # This helps handle slight differences in rounding or conversions.
        tolerance_base = max(extracted_val_nm * 0.05, 1.0)

        # 2) For each row in DB, convert its max_torque to Nm if needed
        for i, row in enumerate(table_data):
            db_torque = row["max_torque"]
            db_unit_lower = row["unit"].lower().strip()

            # Convert DB torque to Nm
            if db_unit_lower in FT_LB_SYNONYMS:
                db_torque_nm = ftlb_to_nm(db_torque)
            elif db_unit_lower in IN_LB_SYNONYMS:
                db_torque_nm = inlb_to_nm(db_torque)
            elif db_unit_lower in NM_SYNONYMS:
                db_torque_nm = db_torque
            else:
                # If the DB unit is unknown, treat as if it was already Nm
                db_torque_nm = db_torque

            # Compare
            if abs(db_torque_nm - extracted_val_nm) <= tolerance_base:
                # We have a match
                self.max_torque_combo.setCurrentIndex(i)
                self.selected_row = row
                self.display_pre_test_rows()
                return
        # If no match found, do nothing => user can pick manually

    # ---------------- Settings Tab ----------------
    def init_settings_tab(self):
        self.settings_tab = QWidget()
        layout = QVBoxLayout(self.settings_tab)

        self.settings_combo = QComboBox()
        self.settings_combo.addItem("Data Management")
        self.settings_combo.addItem("OpenAI Settings")
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

        # New: Checkbox to show/hide extracted data in the Testing Tab
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
        self.settings_tab.setLayout(layout)
        self.tab_widget.addTab(self.settings_tab, "Settings")

        self.load_torque_table_data()

    def on_settings_combo_changed(self, index):
        self.settings_stacked.setCurrentIndex(index)

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
        # Save all settings to the DB
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
        self.report_templates_tab.setLayout(layout)
        self.tab_widget.addTab(self.report_templates_tab, "Report Templates")

    def open_template_editor(self):
        from template_editor import TemplateEditor
        self.template_editor = TemplateEditor()
        self.template_editor.show()

def main():
    # This function can be used as an entry point if needed.
    pass
