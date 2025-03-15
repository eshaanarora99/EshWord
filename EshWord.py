import sys
import json
import html2text
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTextEdit, QFileDialog, QFontDialog,
    QToolBar, QMessageBox, QTabWidget, QDockWidget, QWidget, QVBoxLayout,
    QPushButton, QCheckBox
)
from PyQt6.QtGui import QAction, QIcon, QFont, QTextCharFormat, QColor, QSyntaxHighlighter
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog
from PyQt6.QtCore import QTimer, Qt
from pygments.lexers import guess_lexer

# ---------------------- #
# Syntax Highlighter Code
# ---------------------- #
class CodeHighlighter(QSyntaxHighlighter):
    """
    A basic syntax highlighter that applies blue formatting to common programming keywords.
    This highlighter activates only if the parent QTextEdit widget has its "is_code" property set to True.
    """
    def highlightBlock(self, text):
        if self.parent().property("is_code"):
            keyword_format = QTextCharFormat()
            keyword_format.setForeground(QColor("blue"))
            for word in ["def", "class", "import", "return", "if", "else", "while", "for"]:
                start_index = text.find(word)
                if start_index != -1:
                    self.setFormat(start_index, len(word), keyword_format)

# ----------------------------- #
# Main Application Class: EshWord
# ----------------------------- #
class EshWord(QMainWindow):
    """Main application window for EshWord."""

    def __init__(self):
        super().__init__()

        # Window properties.
        self.setWindowTitle("EshWord - Advanced Word Processor")
        self.setGeometry(100, 100, 1000, 700)
        self.autosave_enabled = True  # Autosave is on by default.
        self.dark_mode = False         # Start in light mode.

        # Create multi-tab editor.
        self.tabs = QTabWidget(self)
        self.tabs.setTabsClosable(True)
        self.tabs.tabCloseRequested.connect(self.close_tab)
        self.setCentralWidget(self.tabs)

        # Create toolbar and sidebar.
        self.create_toolbar()
        self.create_sidebar()

        # Apply initial styling.
        self.apply_styles()

        # Start autosave timer.
        self.enable_autosave()

        # Open a default empty tab.
        self.add_new_tab("Untitled", "")

        self.show()

    # ----------------------------- #
    # File Loading and Saving Methods
    # ----------------------------- #
    def load_file(self):
        """Opens and loads an .esh or .txt file into a new tab."""
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Open File", "",
            "ESH Files (*.esh);;Text Files (*.txt);;All Files (*)"
        )
        if file_name:
            try:
                if file_name.endswith(".esh"):
                    with open(file_name, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        if "text" not in data:
                            raise ValueError("Invalid file format: missing 'text' key.")
                        content = data["text"]
                else:
                    with open(file_name, "r", encoding="utf-8") as f:
                        content = f.read()
                self.add_new_tab(file_name, content)
                self.statusBar().showMessage(f"File loaded: {file_name}", 3000)
            except Exception as e:
                QMessageBox.critical(self, "Load Error", f"Failed to load file: {e}")

    def save_file(self):
        """Saves the currently active document.
        If the file is 'Untitled', it prompts for a file name.
        """
        current_index = self.tabs.currentIndex()
        title = self.tabs.tabText(current_index)
        if title == "Untitled":
            self.save_file_as()
        else:
            self.write_to_file(title)

    def save_file_as(self):
        """Prompts for a file name and saves the current document.
        Supports saving as .esh (JSON with formatting) or plain text (.txt).
        """
        file_name, _ = QFileDialog.getSaveFileName(
            self, "Save File As", "",
            "ESH Files (*.esh);;Text Files (*.txt);;All Files (*)"
        )
        if file_name:
            try:
                if file_name.endswith(".esh"):
                    self.write_to_file(file_name, is_esh=True)
                else:
                    with open(file_name, "w", encoding="utf-8") as f:
                        f.write(self.current_editor().toPlainText())
                current_index = self.tabs.currentIndex()
                self.tabs.setTabText(current_index, file_name)
                self.statusBar().showMessage(f"File saved: {file_name}", 3000)
            except Exception as e:
                QMessageBox.critical(self, "Save Error", f"Failed to save file: {e}")

    def write_to_file(self, file_name, is_esh=True):
        """Writes the current document to file.
        For .esh files, saves in JSON format with formatting info.
        """
        editor = self.current_editor()
        try:
            if is_esh:
                data = {
                    "text": editor.toPlainText(),
                    "font": editor.currentFont().family(),
                    "size": editor.currentFont().pointSize(),
                    "bold": editor.currentFont().bold(),
                    "italic": editor.currentFont().italic(),
                    "underline": editor.currentFont().underline(),
                }
                with open(file_name, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=4)
            else:
                with open(file_name, "w", encoding="utf-8") as f:
                    f.write(editor.toPlainText())
        except Exception as e:
            raise e

    def export_to_pdf(self):
        """Exports the current document to a PDF file."""
        file_name, _ = QFileDialog.getSaveFileName(
            self, "Export to PDF", "",
            "PDF Files (*.pdf);;All Files (*)"
        )
        if file_name:
            try:
                if not file_name.endswith(".pdf"):
                    file_name += ".pdf"
                printer = QPrinter(QPrinter.PrinterMode.HighResolution)
                printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
                printer.setOutputFileName(file_name)
                self.current_editor().document().print(printer)
                self.statusBar().showMessage(f"Exported to PDF: {file_name}", 3000)
            except Exception as e:
                QMessageBox.critical(self, "Export Error", f"Failed to export PDF: {e}")

    def export_to_markdown(self):
        """Exports the current document to a Markdown (.md) file.
        Converts rich text (HTML) to Markdown using html2text.
        """
        file_name, _ = QFileDialog.getSaveFileName(
            self, "Export to Markdown", "",
            "Markdown Files (*.md);;All Files (*)"
        )
        if file_name:
            try:
                html_content = self.current_editor().document().toHtml()
                md_content = html2text.html2text(html_content)
                with open(file_name, "w", encoding="utf-8") as f:
                    f.write(md_content)
                self.statusBar().showMessage(f"Exported to Markdown: {file_name}", 3000)
            except Exception as e:
                QMessageBox.critical(self, "Export Error", f"Failed to export Markdown: {e}")

    def print_document(self):
        """Opens a print dialog and prints the current document."""
        editor = self.current_editor()
        if editor:
            try:
                printer = QPrinter()
                dialog = QPrintDialog(printer, self)
                if dialog.exec():
                    editor.document().print(printer)
            except Exception as e:
                QMessageBox.critical(self, "Print Error", f"Failed to print document: {e}")

    # ----------------------------- #
    # Toolbar and Sidebar Methods
    # ----------------------------- #
    def create_toolbar(self):
        """Creates a toolbar with file, formatting, export, and font adjustment actions."""
        toolbar = QToolBar("Main Toolbar")
        self.addToolBar(toolbar)

        # File actions.
        open_action = QAction(QIcon("icons/open.png"), "Open", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.load_file)
        toolbar.addAction(open_action)

        save_action = QAction(QIcon("icons/save.png"), "Save", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_file)
        toolbar.addAction(save_action)

        print_action = QAction(QIcon("icons/print.png"), "Print", self)
        print_action.setShortcut("Ctrl+P")
        print_action.triggered.connect(self.print_document)
        toolbar.addAction(print_action)

        select_all_action = QAction(QIcon("icons/selectall.png"), "Select All", self)
        select_all_action.setShortcut("Ctrl+A")
        select_all_action.triggered.connect(self.select_all)
        toolbar.addAction(select_all_action)

        font_action = QAction(QIcon("icons/font.png"), "Font", self)
        font_action.setShortcut("Ctrl+T")
        font_action.triggered.connect(self.select_font)
        toolbar.addAction(font_action)

        # Text formatting.
        bold_action = QAction(QIcon("icons/bold.png"), "Bold", self)
        bold_action.setShortcut("Ctrl+B")
        bold_action.triggered.connect(self.make_bold)
        toolbar.addAction(bold_action)

        italic_action = QAction(QIcon("icons/italic.png"), "Italic", self)
        italic_action.setShortcut("Ctrl+I")
        italic_action.triggered.connect(self.make_italic)
        toolbar.addAction(italic_action)

        underline_action = QAction(QIcon("icons/underline.png"), "Underline", self)
        underline_action.setShortcut("Ctrl+U")
        underline_action.triggered.connect(self.make_underline)
        toolbar.addAction(underline_action)

        # Font size adjustments.
        increase_font_action = QAction(QIcon("icons/increase.png"), "Increase Font Size", self)
        increase_font_action.setShortcut("Ctrl++")
        increase_font_action.triggered.connect(self.increase_font_size)
        toolbar.addAction(increase_font_action)

        decrease_font_action = QAction(QIcon("icons/decrease.png"), "Decrease Font Size", self)
        decrease_font_action.setShortcut("Ctrl+-")
        decrease_font_action.triggered.connect(self.decrease_font_size)
        toolbar.addAction(decrease_font_action)

        # Export actions.
        export_pdf_action = QAction(QIcon("icons/pdf.png"), "Export PDF", self)
        export_pdf_action.setShortcut("Ctrl+Shift+P")
        export_pdf_action.triggered.connect(self.export_to_pdf)
        toolbar.addAction(export_pdf_action)

        export_md_action = QAction(QIcon("icons/md.png"), "Export Markdown", self)
        export_md_action.setShortcut("Ctrl+Shift+M")
        export_md_action.triggered.connect(self.export_to_markdown)
        toolbar.addAction(export_md_action)

    def create_sidebar(self):
        """Creates a sidebar with a New Document button and toggle checkboxes (with labels)."""
        self.sidebar = QDockWidget("Options", self)
        self.sidebar.setAllowedAreas(Qt.DockWidgetArea.LeftDockWidgetArea)
        self.sidebar.setMinimumWidth(200)
        container = QWidget()
        layout = QVBoxLayout()

        new_doc_button = QPushButton("New Document")
        new_doc_button.clicked.connect(self.new_document)
        layout.addWidget(new_doc_button)

        self.open_btn = QPushButton("Open File")
        self.open_btn.clicked.connect(self.load_file)
        layout.addWidget(self.open_btn)

        self.save_btn = QPushButton("Save File")
        self.save_btn.clicked.connect(self.save_file)
        layout.addWidget(self.save_btn)

        self.autosave_checkbox = QCheckBox("Enable Autosave")
        self.autosave_checkbox.setChecked(True)
        self.autosave_checkbox.stateChanged.connect(self.toggle_autosave)
        layout.addWidget(self.autosave_checkbox)

        self.darkmode_checkbox = QCheckBox("Enable Dark Mode")
        self.darkmode_checkbox.setChecked(False)
        self.darkmode_checkbox.stateChanged.connect(self.toggle_dark_mode)
        layout.addWidget(self.darkmode_checkbox)

        container.setLayout(layout)
        self.sidebar.setWidget(container)
        self.addDockWidget(Qt.DockWidgetArea.LeftDockWidgetArea, self.sidebar)

    # ----------------------------- #
    # New Document Method
    # ----------------------------- #
    def new_document(self):
        """Creates a new document tab with an empty QTextEdit."""
        self.add_new_tab("Untitled", "")

    # ----------------------------- #
    # Text Formatting Methods
    # ----------------------------- #
    def select_font(self):
        """Opens a font selection dialog and applies the chosen font to the current editor."""
        editor = self.current_editor()
        if editor:
            font, ok = QFontDialog.getFont()
            if ok:
                editor.setFont(font)

    def make_bold(self):
        """Toggles bold formatting in the current editor."""
        editor = self.current_editor()
        if editor:
            font = editor.currentFont()
            font.setBold(not font.bold())
            editor.setCurrentFont(font)

    def make_italic(self):
        """Toggles italic formatting in the current editor."""
        editor = self.current_editor()
        if editor:
            font = editor.currentFont()
            font.setItalic(not font.italic())
            editor.setCurrentFont(font)

    def make_underline(self):
        """Toggles underline formatting in the current editor."""
        editor = self.current_editor()
        if editor:
            font = editor.currentFont()
            font.setUnderline(not font.underline())
            editor.setCurrentFont(font)

    def select_all(self):
        """Selects all text in the current editor."""
        editor = self.current_editor()
        if editor:
            editor.selectAll()

    
    def increase_font_size(self):
        """Increases the font size of the current editor's text."""
        try:
            editor = self.current_editor()
            font = editor.currentFont()
            size = font.pointSize()
            if size <= 0:
                size = 10  # Default value if current size is invalid.
            font.setPointSize(size + 1)
            editor.setFont(font)
        except Exception as e:
            QMessageBox.critical(self, "Font Error", f"Failed to increase font size: {e}")

    def decrease_font_size(self):
        """Decreases the font size of the current editor's text."""
        try:
            editor = self.current_editor()
            font = editor.currentFont()
            size = font.pointSize()
            if size <= 0:
                size = 10  # Default value if current size is invalid.
            if size > 1:  # Ensure we never set the font size to 0.
                font.setPointSize(size - 1)
                editor.setFont(font)
        except Exception as e:
            QMessageBox.critical(self, "Font Error", f"Failed to decrease font size: {e}")


    # ----------------------------- #
    # Tab and Autosave Methods
    # ----------------------------- #
    def add_new_tab(self, title="Untitled", content=""):
        """Creates a new document tab with a QTextEdit widget."""
        editor = QTextEdit()
        editor.setPlainText(content)
        # Optionally set a default font for each new editor.
        default_font = QFont("Segoe UI", 10)
        editor.setFont(default_font)
        index = self.tabs.addTab(editor, title)
        self.tabs.setCurrentIndex(index)

    def close_tab(self, index):
        """Closes the tab at the given index and ensures at least one tab remains open."""
        self.tabs.removeTab(index)
        if self.tabs.count() == 0:
            self.add_new_tab("Untitled", "")

    def current_editor(self):
        """Returns the currently active QTextEdit widget."""
        return self.tabs.currentWidget()

    def enable_autosave(self):
        """Starts a QTimer to autosave every 60 seconds if enabled."""
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.autosave)
        self.timer.start(60000)  # 60,000 ms = 1 minute

    def autosave(self):
        """Autosaves the current document if autosave is enabled.
        For files with a .esh extension, saves as JSON; otherwise, saves as plain text.
        """
        if self.autosave_enabled:
            editor = self.current_editor()
            title = self.tabs.tabText(self.tabs.currentIndex())
            if title != "Untitled":
                try:
                    if title.endswith(".esh"):
                        self.write_to_file(title, is_esh=True)
                    else:
                        with open(title, "w", encoding="utf-8") as f:
                            f.write(editor.toPlainText())
                    self.statusBar().showMessage(f"Autosaved: {title}", 2000)
                except Exception as e:
                    self.statusBar().showMessage(f"Autosave failed: {e}", 2000)

    # ----------------------------- #
    # Toggle Methods
    # ----------------------------- #
    def toggle_autosave(self, state):
        """Toggles autosave on or off based on the checkbox state."""
        if state == Qt.CheckState.Checked:
            self.autosave_enabled = True
        else:
            self.autosave_enabled = False

    def toggle_dark_mode(self, state):
        """Toggles dark mode on or off and updates the UI styling."""
        if state == Qt.CheckState.Checked:
            self.dark_mode = True
        else:
            self.dark_mode = False
        self.apply_styles()

    def apply_styles(self):
        """Applies dark or light mode styling to the application, including hover effects and checkbox label colors."""
        common_font = "Segoe UI"
        if self.dark_mode:
            style = f"""
            QMainWindow {{ background-color: #2C2C2C; font-family: '{common_font}'; }}
            QTextEdit {{ background-color: #3C3C3C; color: #E0E0E0; font-size: 14px; font-family: '{common_font}'; }}
            QToolBar {{ background: #444444; }}
            QToolButton {{ background: transparent; }}
            QToolButton:hover {{ background-color: rgba(255, 255, 255, 0.1); }}
            QDockWidget {{ background: #333333; color: #E0E0E0; }}
            QPushButton {{ background-color: #555555; color: #E0E0E0; font-family: '{common_font}'; }}
            QPushButton:hover {{ background-color: #666666; }}
            QCheckBox {{ color: #E0E0E0; font-family: '{common_font}'; }}
            """
        else:
            style = f"""
            QMainWindow {{ background-color: #f8f9fa; font-family: '{common_font}'; }}
            QTextEdit {{ background-color: white; color: black; font-size: 14px; font-family: '{common_font}'; }}
            QToolBar {{ background: #333333; }}
            QToolButton {{ background: transparent; }}
            QToolButton:hover {{ background-color: rgba(0, 0, 0, 0.1); }}
            QDockWidget {{ background: #eeeeee; color: black; }}
            QPushButton {{ background-color: #444444; color: white; font-family: '{common_font}'; }}
            QPushButton:hover {{ background-color: #555555; }}
            QCheckBox {{ color: black; font-family: '{common_font}'; }}
            """
        self.setStyleSheet(style)

# ----------------------------- #
# Main Execution Block
# ----------------------------- #
if __name__ == "__main__":
    app = QApplication(sys.argv)
    # Set a default font for the entire application.
    default_font = QFont("Segoe UI", 10)
    app.setFont(default_font)
    window = EshWord()
    sys.exit(app.exec())
