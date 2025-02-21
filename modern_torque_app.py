import os
import re  # For extracting numeric part from max torque text
import json
import threading
import pandas as pd
import serial.tools.list_ports
from PIL import Image
import pytesseract
import openai

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QGridLayout, QLabel,
    QComboBox, QPushButton, QTreeWidget, QTreeWidgetItem, QFileDialog,
    QMessageBox, QStatusBar, QTabWidget, QTableWidget, QTableWidgetItem,
    QHeaderView, QDialog, QFormLayout, QLineEdit, QDialogButtonBox, QHBoxLayout,
    QStackedWidget, QDoubleSpinBox
)
from PyQt6.QtCore import QThread, pyqtSignal

# Import your DB handler (DuckDB-based) and serial reader
from db_handler_local import (
    get_torque_table, insert_raw_data, insert_summary,
    add_torque_entry, update_torque_entry, delete_torque_entry
)
from serial_reader import read_from_serial, find_fits_in_selected_row


# ---------------- Worker Thread for Serial Reading ----------------
class SerialReaderWorker(QThread):
    reading_signal = pyqtSignal(float, list)  # (torque_value, fits_list)

    def __init__(self, port, selected_row):
        super().__init__()
        self.port = port
        self.selected_row = selected_row
        self.stop_event = threading.Event()

    def run(self):
        BAUD_RATE = 9600

        def callback(target_torque):
            # Debug: print the torque value as it arrives
            print(f"[DEBUG] Serial callback received torque: {target_torque}")
            if self.stop_event.is_set():
                return
            fits = find_fits_in_selected_row(target_torque, self.selected_row)
            if fits:
                print(f"[DEBUG] torque {target_torque} fits in ranges: {fits}")
            else:
                print(f"[DEBUG] torque {target_torque} did NOT fit any range")
            # Always emit the reading
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
    """
    Multiply max_torque by 0.916, 0.583, and 0.333,
    rounding each result to the nearest 10.
    Returns a list of three numeric values.
    """
    factors = [0.916, 0.583, 0.333]
    results = []
    for f in factors:
        raw_val = max_torque * f
        rounded = round(raw_val / 10) * 10
        results.append(rounded)
    print(f"[DEBUG] calc_applied_torques({max_torque}) => {results}")
    return results

def calc_allowance_range(applied_val: float) -> str:
    """
    Calculate the allowance range.
      - Use ±6% if applied_val < 10,
      - Otherwise, use ±4%.
    Round each boundary to one decimal.
    """
    if applied_val < 10:
        tolerance = 0.06
    else:
        tolerance = 0.04
    low = applied_val * (1 - tolerance)
    high = applied_val * (1 + tolerance)
    range_str = f"{round(low,1)} - {round(high,1)}"
    print(f"[DEBUG] calc_allowance_range({applied_val}) => {range_str}")
    return range_str


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

        # When Max Torque changes, extract numeric part and update applied torques and allowances.
        self.max_torque_edit.textChanged.connect(self.auto_fill_applied_from_max)
        # If user edits the applied torque JSON directly, recalc allowances.
        self.applied_torq_edit.textChanged.connect(self.auto_fill_allowances_from_applied)

        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

    def auto_fill_applied_from_max(self):
        txt = self.max_torque_edit.text().strip()
        if not txt:
            print("[DEBUG] Max Torque field is empty.")
            return
        match = re.search(r"[\d\.]+", txt)
        if match:
            try:
                max_torque = float(match.group())
                print(f"[DEBUG] Extracted numeric max_torque from '{txt}' => {max_torque}")
            except ValueError:
                print(f"[DEBUG] Could not convert '{match.group()}' to float.")
                return
        else:
            print(f"[DEBUG] No numeric portion found in '{txt}'")
            return

        applied_list = calc_applied_torques(max_torque)
        self.applied_torq_edit.blockSignals(True)
        self.applied_torq_edit.setText(json.dumps(applied_list))
        self.applied_torq_edit.blockSignals(False)

        self.auto_fill_allowances_from_applied()

    def auto_fill_allowances_from_applied(self):
        txt = self.applied_torq_edit.text().strip()
        if not txt:
            print("[DEBUG] Applied Torque (JSON) field is empty.")
            return
        try:
            arr = json.loads(txt)
            if not isinstance(arr, list):
                print(f"[DEBUG] '{txt}' is not a JSON list.")
                return
        except (ValueError, json.JSONDecodeError):
            print(f"[DEBUG] Could not parse JSON from '{txt}'.")
            return

        for i in range(3):
            val = arr[i] if i < len(arr) else 0
            allow_str = calc_allowance_range(val)
            if i == 0:
                self.allowance1_edit.setText(allow_str)
            elif i == 1:
                self.allowance2_edit.setText(allow_str)
            else:
                self.allowance3_edit.setText(allow_str)

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
        self.setStyleSheet(self.load_stylesheet())

        # For storing test results, customer info, and the active row
        self.results_by_range = {}
        self.customer_info = {}
        self.serial_worker = None
        self.selected_row = None

        # ------------- OpenAI Settings -------------
        # API key and model config
        self.openai_api_key = None
        self.openai_model = "gpt-4-turbo"  # default
        self.openai_temperature = 0.7
        self.openai_top_p = 1.0
        self.openai_presence_penalty = 0.0
        self.openai_frequency_penalty = 0.0

        self.init_ui()

    def load_stylesheet(self):
        stylesheet = """
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
        QDialog QLineEdit {
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
        QComboBox, QLineEdit {
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
        QTreeWidget {
            background-color: #FFFFFF;
            border: 1px solid #ccc;
            color: #333;
        }
        QTreeWidget::item {
            padding: 6px;
        }
        QTreeWidget::item:selected {
            background-color: #3498db;
            color: #FFFFFF;
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
        return stylesheet

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
        layout = QVBoxLayout(self.testing_tab)

        form_layout = QGridLayout()
        label_torque = QLabel("Max Torque:")
        self.torque_combo = QComboBox()
        self.torque_combo.currentIndexChanged.connect(self.on_torque_combo_selected)
        form_layout.addWidget(label_torque, 0, 0)
        form_layout.addWidget(self.torque_combo, 0, 1)

        label_port = QLabel("Serial Port:")
        self.port_combo = QComboBox()
        self.port_combo.addItems(self.get_serial_ports())
        form_layout.addWidget(label_port, 1, 0)
        form_layout.addWidget(self.port_combo, 1, 1)

        self.export_excel_btn = QPushButton("Export Summary")
        self.export_excel_btn.clicked.connect(self.export_summary_to_excel)
        form_layout.addWidget(self.export_excel_btn, 2, 0, 1, 2)

        self.upload_info_btn = QPushButton("Import Customer Info")
        self.upload_info_btn.clicked.connect(self.upload_customer_info)
        form_layout.addWidget(self.upload_info_btn, 3, 0, 1, 2)

        self.start_btn = QPushButton("Begin Test")
        self.start_btn.clicked.connect(self.start_test)
        form_layout.addWidget(self.start_btn, 4, 0)

        self.stop_btn = QPushButton("End Test")
        self.stop_btn.clicked.connect(self.stop_test)
        self.stop_btn.setEnabled(False)
        form_layout.addWidget(self.stop_btn, 4, 1)

        layout.addLayout(form_layout)

        # Add a live torque label to display the current reading
        self.live_torque_label = QLabel("Live Torque: --")
        self.live_torque_label.setStyleSheet("font-size: 16px; padding: 5px;")
        layout.addWidget(self.live_torque_label)

        self.tree = QTreeWidget()
        self.tree.setColumnCount(7)
        self.tree.setHeaderLabels([
            "Applied Torque", "Allowance", "Test 1", "Test 2",
            "Test 3", "Test 4", "Test 5"
        ])
        self.tree.header().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.tree)

        self.tab_widget.addTab(self.testing_tab, "Torque Testing")
        self.refresh_torque_dropdown()

    def get_serial_ports(self):
        ports = serial.tools.list_ports.comports()
        return [port.device for port in ports]

    def refresh_torque_dropdown(self):
        self.torque_table = get_torque_table()
        self.torque_combo.clear()
        for row in self.torque_table:
            display_text = f"{row['max_torque']} {row['unit']} - {row['type']}"
            self.torque_combo.addItem(display_text)
        if self.torque_table:
            self.torque_combo.setCurrentIndex(0)
            self.selected_row = self.torque_table[0]
            self.display_pre_test_rows()
        else:
            self.selected_row = None

    def on_torque_combo_selected(self, index):
        if index < 0 or index >= len(self.torque_table):
            return
        self.selected_row = self.torque_table[index]
        self.results_by_range = {}
        print(f"[DEBUG] on_torque_combo_selected => selected_row: {self.selected_row}")
        self.display_pre_test_rows()

    def display_pre_test_rows(self):
        self.tree.clear()
        if not self.selected_row:
            print("[DEBUG] No selected row in display_pre_test_rows.")
            return
        try:
            applied_arr = json.loads(self.selected_row.get("applied_torq", "[]"))
        except json.JSONDecodeError:
            applied_arr = [0, 0, 0]
        print(f"[DEBUG] display_pre_test_rows => applied_torq array: {applied_arr}")
        for i in range(3):
            allowance_key = self.selected_row.get(f"allowance{i+1}", "")
            applied_val = applied_arr[i] if i < len(applied_arr) else 0
            row_values = [str(applied_val), allowance_key] + [""] * 5
            item = QTreeWidgetItem(row_values)
            self.tree.addTopLevelItem(item)

    def start_test(self):
        if not self.port_combo.currentText():
            QMessageBox.critical(self, "Error", "No serial port selected.")
            return
        if not self.torque_combo.currentText():
            QMessageBox.critical(self, "Error", "No torque entry selected.")
            return

        self.selected_row = self.torque_table[self.torque_combo.currentIndex()]
        self.results_by_range = {}
        print("[DEBUG] start_test => selected row:", self.selected_row)
        self.display_pre_test_rows()

        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.statusBar.showMessage("Test in progress...")
        print("[DEBUG] Test started.")

        port = self.port_combo.currentText()
        self.serial_worker = SerialReaderWorker(port, self.selected_row)
        self.serial_worker.reading_signal.connect(self.process_reading)
        self.serial_worker.start()

    def stop_test(self):
        print("[DEBUG] stop_test called.")
        if self.serial_worker:
            self.serial_worker.stop()
            self.serial_worker.wait(2000)

        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.update_summary_table()
        self.save_final_summary()

        self.statusBar.showMessage("Test ended. Summary updated.")
        QMessageBox.information(self, "Test Completed", "Test ended and summary updated.")
        print("[DEBUG] Test stopped. Results by range =>", self.results_by_range)

    def process_reading(self, target_torque, fits):
        # Update the live torque cell with the current reading
        self.live_torque_label.setText(f"Live Torque: {target_torque}")
        if fits:
            self.live_torque_label.setStyleSheet("background-color: green; color: white; font-size: 16px; padding: 5px;")
        else:
            self.live_torque_label.setStyleSheet("background-color: red; color: white; font-size: 16px; padding: 5px;")

        # Process each fit (if any) for summary updates
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
        print("[DEBUG] update_summary_table => results_by_range:", self.results_by_range)
        self.tree.clear()
        if not self.selected_row:
            return
        try:
            applied_arr = json.loads(self.selected_row.get("applied_torq", "[]"))
        except json.JSONDecodeError:
            applied_arr = [0, 0, 0]
        for i in range(3):
            allowance_key = self.selected_row.get(f"allowance{i+1}", "")
            applied_val = applied_arr[i] if i < len(applied_arr) else 0.0
            test_values = self.results_by_range.get(allowance_key, [])
            row_values = [str(applied_val), allowance_key] + [str(v) for v in test_values]
            while len(row_values) < 7:
                row_values.append("")
            item = QTreeWidgetItem(row_values)
            self.tree.addTopLevelItem(item)
            actual_numbers = [v for v in test_values if isinstance(v, (float, int))]
            insert_summary(allowance_key, actual_numbers)

    def save_final_summary(self):
        print("[DEBUG] save_final_summary => Called. (placeholder logic)")
        print("Final summary saved to database (placeholder).")

    def export_summary_to_excel(self):
        print("[DEBUG] export_summary_to_excel => Called.")
        summary_data = []
        for allow_range, results in self.results_by_range.items():
            valid_results = [r for r in results if isinstance(r, (float, int))]
            while len(valid_results) < 5:
                valid_results.append("")
            summary_data.append({
                "Allowance Range": allow_range,
                "Test 1": valid_results[0],
                "Test 2": valid_results[1],
                "Test 3": valid_results[2],
                "Test 4": valid_results[3],
                "Test 5": valid_results[4],
            })
        if summary_data:
            df = pd.DataFrame(summary_data)
            df['LowerBound'] = df['Allowance Range'].apply(lambda x: float(x.split('-')[0].strip()) if '-' in x else 0)
            df = df.sort_values(by='LowerBound', ascending=False)
            df.drop(columns=['LowerBound'], inplace=True)
            try:
                df.to_excel("summary.xlsx", index=False)
                QMessageBox.information(self, "Export Summary", "Summary exported successfully to summary.xlsx")
                print("[DEBUG] Summary exported to summary.xlsx.")
            except Exception as e:
                QMessageBox.critical(self, "Export Error", f"Error exporting to Excel:\n{e}")
                print("[DEBUG] Error exporting to Excel:", e)
        else:
            QMessageBox.warning(self, "Export Warning", "No summary data available to export.")
            print("[DEBUG] No summary data available to export.")

    def upload_customer_info(self):
        print("[DEBUG] upload_customer_info => Called.")
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Image", "",
            "Image Files (*.png *.jpg *.jpeg *.bmp);;All Files (*)"
        )
        if not file_path:
            print("[DEBUG] No file selected for OCR.")
            return

        # Ask user if they want to use GPT-4 Vision or Tesseract
        use_openai = QMessageBox.question(
            self, "Use GPT-4 Vision OCR?",
            "Do you want to use OpenAI GPT-4 Vision (requires special access)?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        ) == QMessageBox.StandardButton.Yes

        if use_openai:
            if not self.openai_api_key:
                QMessageBox.critical(self, "Error", "OpenAI API Key not set.")
                return
            # Attempt hypothetical GPT-4 Vision call
            ocr_text = self.ocr_with_openai_vision(file_path)
            if not ocr_text:
                QMessageBox.warning(self, "OpenAI OCR Failed", "No text extracted or an error occurred.")
                return
        else:
            # Local Tesseract fallback
            try:
                image = Image.open(file_path)
                ocr_text = pytesseract.image_to_string(image)
                print("[DEBUG] OCR Text:", ocr_text)
            except Exception as e:
                QMessageBox.critical(self, "OCR Error", f"Error processing image: {e}")
                print("[DEBUG] Error processing image for OCR:", e)
                return

        self.parse_ocr_text(ocr_text)

    def parse_ocr_text(self, ocr_text: str):
        self.customer_info = {}
        for line in ocr_text.splitlines():
            line = line.strip()
            if not line:
                continue
            parts = line.split(":", 1)
            if len(parts) == 2:
                key = parts[0].strip().lower()
                value = parts[1].strip()
                if "customer" in key:
                    self.customer_info["customer"] = value
                elif "email" in key:
                    self.customer_info["email"] = value
                elif "contact" in key:
                    self.customer_info["contact"] = value
                elif "brand" in key:
                    self.customer_info["brand"] = value
                elif "model" in key:
                    self.customer_info["model"] = value
                elif "unit" in key:
                    self.customer_info["unit"] = value
                elif "serial" in key:
                    self.customer_info["serial"] = value
        if self.customer_info:
            self.statusBar.showMessage("Customer info imported.")
            print("[DEBUG] Customer Info parsed =>", self.customer_info)
        else:
            self.statusBar.showMessage("No recognizable customer info found.")
            print("[DEBUG] No recognizable customer info found in OCR text.")

    def ocr_with_openai_vision(self, image_path: str) -> str:
        """
        Hypothetical function for GPT-4 Vision OCR.
        This won't work unless you have special access to a real image endpoint from OpenAI.
        We'll show how you might pass 'model' or 'temperature' if it existed.
        """
        openai.api_key = self.openai_api_key
        try:
            # This is NOT an actual method for images in the openai package
            # It's just a placeholder to illustrate usage.
            response = openai.Image.create(
                image=open(image_path, "rb").read(),
                model=self.openai_model
            )
            recognized_text = response.get("data", {}).get("text", "")
            print("[DEBUG] GPT-4 Vision recognized text:", recognized_text)
            return recognized_text
        except Exception as e:
            print("[DEBUG] OpenAI Vision OCR error:", e)
            return ""

    # ---------------- Settings Tab (with advanced OpenAI settings) ----------------
    def init_settings_tab(self):
        self.settings_tab = QWidget()
        layout = QVBoxLayout(self.settings_tab)

        # A dropdown to pick which sub-section we want
        self.settings_combo = QComboBox()
        self.settings_combo.addItem("Data Management")
        self.settings_combo.addItem("OpenAI Settings")
        self.settings_combo.currentIndexChanged.connect(self.on_settings_combo_changed)
        layout.addWidget(self.settings_combo)

        self.settings_stacked = QStackedWidget()
        layout.addWidget(self.settings_stacked)

        # --- Data Management Page ---
        self.data_management_page = QWidget()
        dm_layout = QVBoxLayout(self.data_management_page)

        self.torque_table_widget = QTableWidget()
        self.torque_table_widget.setColumnCount(4)
        self.torque_table_widget.setHorizontalHeaderLabels([
            "Max Torque", "Unit", "Type", "Applied Torque"
        ])
        self.torque_table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        dm_layout.addWidget(self.torque_table_widget)

        button_layout = QHBoxLayout()
        self.add_btn = QPushButton("Add Entry")
        self.add_btn.clicked.connect(self.add_entry)
        self.edit_btn = QPushButton("Edit Entry")
        self.edit_btn.clicked.connect(self.edit_entry)
        self.delete_btn = QPushButton("Delete Entry")
        self.delete_btn.clicked.connect(self.delete_entry)
        self.refresh_btn = QPushButton("Refresh")
        self.refresh_btn.clicked.connect(self.load_torque_table_data)
        button_layout.addWidget(self.add_btn)
        button_layout.addWidget(self.edit_btn)
        button_layout.addWidget(self.delete_btn)
        button_layout.addWidget(self.refresh_btn)
        dm_layout.addLayout(button_layout)

        self.data_management_page.setLayout(dm_layout)

        # --- OpenAI Settings Page ---
        self.openai_settings_page = QWidget()
        openai_layout = QFormLayout(self.openai_settings_page)

        self.api_key_edit = QLineEdit()
        self.api_key_edit.setText(self.openai_api_key if self.openai_api_key else "")
        openai_layout.addRow("OpenAI API Key:", self.api_key_edit)

        # Updated model combo with the requested items
        self.model_combo = QComboBox()
        self.model_combo.addItems([
            "o1",
            "gpt-4o",
            "gpt-4o-mini",
            "gpt-4-turbo",
            "o3",
            "o3-mini"
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

        self.settings_stacked.addWidget(self.data_management_page)   # index 0
        self.settings_stacked.addWidget(self.openai_settings_page)   # index 1

        self.settings_tab.setLayout(layout)
        self.tab_widget.addTab(self.settings_tab, "Settings")

        # Finally, load the table data
        self.load_torque_table_data()

    def on_settings_combo_changed(self, index):
        self.settings_stacked.setCurrentIndex(index)

    def save_openai_settings(self):
        self.openai_api_key = self.api_key_edit.text().strip()
        self.openai_model = self.model_combo.currentText()
        self.openai_temperature = self.temp_spin.value()
        self.openai_top_p = self.top_p_spin.value()
        self.openai_presence_penalty = self.presence_spin.value()
        self.openai_frequency_penalty = self.freq_spin.value()

        QMessageBox.information(self, "OpenAI Settings", "OpenAI settings saved in memory.")

    # ---------------- Data Management Methods ----------------
    def load_torque_table_data(self):
        print("[DEBUG] load_torque_table_data => Called.")
        table_data = get_torque_table()
        self.torque_table_widget.setRowCount(len(table_data))
        for i, row in enumerate(table_data):
            self.torque_table_widget.setItem(i, 0, QTableWidgetItem(str(row.get("max_torque", ""))))
            self.torque_table_widget.setItem(i, 1, QTableWidgetItem(row.get("unit", "")))
            self.torque_table_widget.setItem(i, 2, QTableWidgetItem(row.get("type", "")))
            self.torque_table_widget.setItem(i, 3, QTableWidgetItem(row.get("applied_torq", "")))
        self.torque_table_widget.resizeColumnsToContents()
        print("[DEBUG] Torque table loaded with", len(table_data), "rows.")

    def add_entry(self):
        print("[DEBUG] add_entry => Called.")
        dialog = TorqueEntryDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.get_data()
            print("[DEBUG] add_entry => dialog returned data:", data)
            add_torque_entry(
                data["max_torque"], data["unit"], data["type"],
                data["applied_torq"], data["allowance1"],
                data["allowance2"], data["allowance3"]
            )
            QMessageBox.information(self, "Success", "Entry added successfully.")
            self.load_torque_table_data()
            self.refresh_torque_dropdown()

    def edit_entry(self):
        print("[DEBUG] edit_entry => Called.")
        selected_items = self.torque_table_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select an entry to edit.")
            print("[DEBUG] No selection for editing.")
            return
        row = self.torque_table_widget.currentRow()
        table_data = get_torque_table()
        entry = table_data[row]
        print("[DEBUG] edit_entry => current entry:", entry)
        dialog = TorqueEntryDialog(self, entry)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.get_data()
            print("[DEBUG] edit_entry => updated data:", data)
            update_torque_entry(
                entry["id"], data["max_torque"], data["unit"], data["type"],
                data["applied_torq"], data["allowance1"], data["allowance2"], data["allowance3"]
            )
            QMessageBox.information(self, "Success", "Entry updated successfully.")
            self.load_torque_table_data()
            self.refresh_torque_dropdown()

    def delete_entry(self):
        print("[DEBUG] delete_entry => Called.")
        selected_items = self.torque_table_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select an entry to delete.")
            print("[DEBUG] No selection for deleting.")
            return
        row = self.torque_table_widget.currentRow()
        table_data = get_torque_table()
        entry = table_data[row]
        confirm = QMessageBox.question(
            self, "Confirm Deletion", f"Are you sure you want to delete entry ID {entry['id']}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if confirm == QMessageBox.StandardButton.Yes:
            print(f"[DEBUG] Deleting entry ID {entry['id']}")
            delete_torque_entry(entry["id"])
            QMessageBox.information(self, "Success", "Entry deleted successfully.")
            self.load_torque_table_data()
            self.refresh_torque_dropdown()

    # ---------------- Report Templates Tab ----------------
    def init_report_templates_tab(self):
        self.report_templates_tab = QWidget()
        layout = QVBoxLayout(self.report_templates_tab)

        self.template_editor_btn = QPushButton("Open Report Template Editor")
        self.template_editor_btn.clicked.connect(self.open_template_editor)
        layout.addWidget(self.template_editor_btn)

        self.report_templates_tab.setLayout(layout)
        self.tab_widget.addTab(self.report_templates_tab, "Report Templates")

    def open_template_editor(self):
        print("[DEBUG] open_template_editor => Called.")
        from template_editor import TemplateEditor
        self.template_editor = TemplateEditor()
        self.template_editor.show()
