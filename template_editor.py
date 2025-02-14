import os
from PyQt6.QtWidgets import QMainWindow, QFileDialog
from PyQt6.QtCore import QObject, pyqtSlot, QUrl
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebChannel import QWebChannel

class EditorAPI(QObject):
    @pyqtSlot(result=str)
    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(
            None, "Open HTML File", "", 
            "HTML Files (*.html);;All Files (*)"
        )
        if path:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    return f.read()
            except Exception:
                return ""
        return ""
    
    @pyqtSlot(str, result=str)
    def save_template(self, content):
        try:
            with open("template_saved.html", "w", encoding="utf-8") as f:
                f.write(content)
            return "Template saved successfully."
        except Exception as e:
            return f"Error saving template: {e}"
    
    @pyqtSlot(str, result=str)
    def save_template_as(self, content):
        path, _ = QFileDialog.getSaveFileName(
            None, "Save HTML File As", "", 
            "HTML Files (*.html);;All Files (*)"
        )
        if path:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(content)
                return f"Template saved as {path}"
            except Exception as e:
                return f"Error saving template: {e}"
        return "Save cancelled."

class TemplateEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Template Editor")
        self.setGeometry(150, 150, 800, 600)
        
        self.web_view = QWebEngineView(self)
        self.setCentralWidget(self.web_view)
        
        self.channel = QWebChannel(self.web_view.page())
        self.editor_api = EditorAPI()
        self.channel.registerObject("editorAPI", self.editor_api)
        self.web_view.page().setWebChannel(self.channel)
        
        local_file = os.path.join(os.path.dirname(__file__), "editor.html")
        self.web_view.load(QUrl.fromLocalFile(local_file))
