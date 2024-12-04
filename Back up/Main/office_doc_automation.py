import os
import json
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog,
    QListWidget, QMessageBox, QInputDialog, QScrollArea, QSizePolicy
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt
from docx import Document
import pandas as pd


def create_file_automation_window():
    app = QApplication.instance() or QApplication([])

    # Create the main window
    file_automation_window = QWidget()
    file_automation_window.setWindowTitle("File Automation System")
    file_automation_window.setMinimumSize(900, 700)
    file_automation_window.setStyleSheet("background-color: #f4f4f4;")

    # Create the layout and add components
    layout = QVBoxLayout()

    # Title label
    title_label = QLabel("File Automation System")
    title_label.setStyleSheet("font-size: 22px; font-weight: bold; color: #8b0000;")
    layout.addWidget(title_label, alignment=Qt.AlignCenter)

    # Labels for displaying selected file paths
    template_path_label = QLabel("No template selected.")
    excel_path_label = QLabel("No Excel file selected.")
    save_path_label = QLabel("No save directory selected.")
    for label in [template_path_label, excel_path_label, save_path_label]:
        label.setStyleSheet("font-size: 14px; color: #333;")
        layout.addWidget(label)

    # List to display saved automations with scrolling
    automation_list = QListWidget()
    automation_list.setStyleSheet("""
        font-size: 16px;
        background-color: #ffffff;
        border: 1px solid #ccc;
        padding: 5px;
    """)
    automation_list.setFixedHeight(150)
    automation_list.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
    layout.addWidget(automation_list)

    # Buttons for file selection and saving
    browse_template_button = QPushButton("Select Word Template")
    browse_excel_button = QPushButton("Select Excel File")
    browse_save_button = QPushButton("Select Save Directory")
    save_automation_button = QPushButton("Save Automation")
    delete_automation_button = QPushButton("Delete Selected Automation")
    run_automation_button = QPushButton("Run Automation")

    button_style = """
        QPushButton {
            font-size: 14px;
            color: white;
            background-color: #8b0000;
            border: none;
            padding: 10px;
            margin: 5px 0;
        }
        QPushButton:hover {
            background-color: #a00000;
        }
    """
    for button in [browse_template_button, browse_excel_button, browse_save_button, save_automation_button, delete_automation_button, run_automation_button]:
        button.setStyleSheet(button_style)

    # Add buttons to the layout
    layout.addWidget(browse_template_button)
    layout.addWidget(browse_excel_button)
    layout.addWidget(browse_save_button)
    layout.addWidget(save_automation_button)
    layout.addWidget(delete_automation_button)
    layout.addWidget(run_automation_button)

    # Variables to hold selected file paths
    selected_template_path = None
    selected_excel_path = None
    save_directory_path = None
    automations = {}

    # Function to browse and select a Word template
    def browse_template():
        nonlocal selected_template_path
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(
            file_automation_window, "Select Word Template", "", "Word Files (*.docx);;All Files (*)", options=options
        )
        if file_name:
            template_path_label.setText(f"Selected Template: {file_name}")
            selected_template_path = file_name
        else:
            template_path_label.setText("No template selected.")
            selected_template_path = None

    # Function to browse and select an Excel file
    def browse_excel():
        nonlocal selected_excel_path
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(
            file_automation_window, "Select Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options
        )
        if file_name:
            excel_path_label.setText(f"Selected Excel File: {file_name}")
            selected_excel_path = file_name
        else:
            excel_path_label.setText("No Excel file selected.")
            selected_excel_path = None

    # Function to browse and select a save directory
    def browse_save_directory():
        nonlocal save_directory_path
        directory = QFileDialog.getExistingDirectory(
            file_automation_window, "Select Save Directory", options=QFileDialog.ShowDirsOnly
        )
        if directory:
            save_path_label.setText(f"Save Directory: {directory}")
            save_directory_path = directory
        else:
            save_path_label.setText("No directory selected.")
            save_directory_path = None

    # Function to save automation settings
    def save_automation():
        if not (selected_template_path and selected_excel_path and save_directory_path):
            QMessageBox.warning(
                file_automation_window, "Warning", "Please select all required files and a save directory."
            )
            return

        # Prompt the user to name the automation
        automation_name, ok = QInputDialog.getText(
            file_automation_window, "Name Automation", "Enter a name for the automation:"
        )
        if not ok or not automation_name.strip():
            QMessageBox.warning(file_automation_window, "Warning", "Automation name cannot be empty.")
            return

        # Check for duplicate names
        if automation_name in automations:
            QMessageBox.warning(file_automation_window, "Warning", "An automation with this name already exists.")
            return

        # Save automation configuration
        automations[automation_name] = {
            "template": selected_template_path,
            "excel": selected_excel_path,
            "save_path": save_directory_path
        }

        # Save to file
        save_automations()
        update_automation_list()
        QMessageBox.information(file_automation_window, "Success", f"Automation '{automation_name}' saved successfully!")

    # Function to delete selected automation
    def delete_selected_automation():
        selected_item = automation_list.currentItem()
        if not selected_item:
            QMessageBox.warning(file_automation_window, "Warning", "Please select an automation to delete.")
            return

        automation_name = selected_item.text()
        del automations[automation_name]
        save_automations()
        update_automation_list()
        QMessageBox.information(file_automation_window, "Success", f"Automation '{automation_name}' deleted successfully.")

    # Function to update the automation list displayed in the UI
    def update_automation_list():
        automation_list.clear()
        for name in automations:
            automation_list.addItem(name)

    # Function to save automations to a JSON file
    def save_automations():
        with open("automations.json", "w") as f:
            json.dump(automations, f)

    # Function to load automations from a JSON file
    def load_automations():
        nonlocal automations
        if os.path.exists("automations.json"):
            with open("automations.json", "r") as f:
                automations = json.load(f)
            update_automation_list()

    # Function to run the selected automation
    def run_automation():
        selected_item = automation_list.currentItem()
        if not selected_item:
            QMessageBox.warning(file_automation_window, "Warning", "Please select an automation to run.")
            return

        automation = automations[selected_item.text()]
        template_path = automation["template"]
        excel_path = automation["excel"]
        save_path = automation["save_path"]

        try:
            # Read Excel and process each row
            data = pd.read_excel(excel_path)
            for idx, row in data.iterrows():
                # Open the template and replace placeholders with actual values
                doc = Document(template_path)
                for paragraph in doc.paragraphs:
                    for column_name, value in row.items():
                        placeholder = f"[{column_name}]"
                        if placeholder in paragraph.text:
                            paragraph.text = paragraph.text.replace(placeholder, str(value))

                # Extract Name and Surname from the current row
                name = row.get("NAME", f"Document_{idx+1}")
                surname = row.get("SURNAME", "")
                filename = f"{name}_{surname}.docx".strip("_")  # Ensure no stray underscores

                # Save the modified document
                output_path = os.path.join(save_path, filename)
                doc.save(output_path)

            QMessageBox.information(file_automation_window, "Success", f"Automation completed. Files saved to {save_path}")
        except Exception as e:
            QMessageBox.critical(file_automation_window, "Error", f"An error occurred: {str(e)}")

    # Connect buttons to functions
    browse_template_button.clicked.connect(browse_template)
    browse_excel_button.clicked.connect(browse_excel)
    browse_save_button.clicked.connect(browse_save_directory)
    save_automation_button.clicked.connect(save_automation)
    delete_automation_button.clicked.connect(delete_selected_automation)
    run_automation_button.clicked.connect(run_automation)

    # Load automations on startup
    load_automations()

    file_automation_window.setLayout(layout)
    file_automation_window.show()
    app.exec_()


if __name__ == "__main__":
    create_file_automation_window()
