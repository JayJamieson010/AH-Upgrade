from PyQt5.QtWidgets import (
    QWidget, QLabel, QVBoxLayout, QPushButton, QTextEdit, QMessageBox, QSplitter, QFileDialog, QApplication
)
from PyQt5.QtCore import Qt
import io
import contextlib

# Global variable to keep the window reference
editor_window = None

def open_editor():
    """Function to create and display the Live Code Editor window."""
    global editor_window  # Use a global variable to hold the window reference

    # Create a new widget as a window
    editor_window = QWidget()
    editor_window.setWindowTitle("Live Code Editor")
    editor_window.setMinimumSize(800, 600)

    # Variables to hold components
    code_editor = QTextEdit()
    output_console = QTextEdit()

    # Function to execute code from the editor
    def run_code():
        """Executes the Python code written in the editor."""
        code = code_editor.toPlainText().strip()
        if not code:
            QMessageBox.warning(editor_window, "Warning", "No code to execute!")
            return

        output_buffer = io.StringIO()
        with contextlib.redirect_stdout(output_buffer), contextlib.redirect_stderr(output_buffer):
            try:
                exec(code)
            except Exception as e:
                output_buffer.write(f"[ERROR] {e}\n")

        output_console.setPlainText(output_buffer.getvalue() or "[INFO] Code executed successfully!")

    # Function to save code to a file
    def save_code():
        """Saves the Python code to a file."""
        code = code_editor.toPlainText()
        if not code.strip():
            QMessageBox.warning(editor_window, "Warning", "No code to save!")
            return

        file_path, _ = QFileDialog.getSaveFileName(editor_window, "Save Code", "", "Python Files (*.py)")
        if file_path:
            try:
                with open(file_path, "w") as file:
                    file.write(code)
                QMessageBox.information(editor_window, "Success", "Code saved successfully!")
            except Exception as e:
                QMessageBox.critical(editor_window, "Error", f"Failed to save code: {e}")

    # Function to load code from a file
    def load_code():
        """Loads Python code from a file."""
        file_path, _ = QFileDialog.getOpenFileName(editor_window, "Open Code", "", "Python Files (*.py)")
        if file_path:
            try:
                with open(file_path, "r") as file:
                    code_editor.setPlainText(file.read())
                QMessageBox.information(editor_window, "Success", "Code loaded successfully!")
            except Exception as e:
                QMessageBox.critical(editor_window, "Error", f"Failed to load code: {e}")

    # Create the layout and components
    layout = QVBoxLayout()

    # Label for instructions
    code_label = QLabel("Write your Python code below:")
    code_label.setStyleSheet("font-size: 16px; font-weight: bold;")
    layout.addWidget(code_label)

    # Splitter for code editor and output console
    splitter = QSplitter(Qt.Vertical)

    # Code editor setup
    code_editor.setPlaceholderText("# Write Python code here...")
    code_editor.setStyleSheet("font-family: Courier; font-size: 16px; line-height: 1.5;")
    splitter.addWidget(code_editor)

    # Output console setup
    output_console.setReadOnly(True)
    output_console.setStyleSheet("background-color: black; color: white; font-family: Courier; font-size: 14px; line-height: 1.5;")
    output_console.setPlaceholderText("Console output will appear here...")
    splitter.addWidget(output_console)

    layout.addWidget(splitter)

    # Buttons for functionality
    run_button = QPushButton("Run Code")
    run_button.setStyleSheet("font-size: 16px; padding: 10px;")
    run_button.clicked.connect(run_code)

    save_button = QPushButton("Save Code")
    save_button.setStyleSheet("font-size: 16px; padding: 10px;")
    save_button.clicked.connect(save_code)

    load_button = QPushButton("Load Code")
    load_button.setStyleSheet("font-size: 16px; padding: 10px;")
    load_button.clicked.connect(load_code)

    # Add buttons to the layout
    layout.addWidget(run_button)
    layout.addWidget(save_button)
    layout.addWidget(load_button)

    # Set the layout and show the window
    editor_window.setLayout(layout)
    editor_window.show()

# Run the application
if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    open_editor()
    sys.exit(app.exec_())
