# file: openai_settings_dialog.py
from PyQt6.QtWidgets import QDialog, QFormLayout, QLineEdit, QDialogButtonBox

class OpenAISettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("OpenAI Settings")
        self.api_key_edit = QLineEdit()
        self.init_ui()

    def init_ui(self):
        layout = QFormLayout(self)
        layout.addRow("OpenAI API Key:", self.api_key_edit)

        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

    def get_api_key(self):
        return self.api_key_edit.text().strip()
