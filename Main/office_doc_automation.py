import os
import json
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog,
    QListWidget, QMessageBox, QInputDialog, QScrollArea, QSizePolicy,  QTextEdit, QHBoxLayout
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt
from docx import Document
import pandas as pd
import time

import os
import pandas as pd
import win32com.client as win32
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton,
    QFileDialog, QMessageBox
)
from PyQt5.QtCore import Qt

bulk_email_window = None

selected_template_path = None
selected_excel_path = None
save_directory_path = None
automations = {}

def create_bulk_email_window():
    global bulk_email_window
    selected_excel_path = None

    def browse_excel():
        nonlocal selected_excel_path
        file_name, _ = QFileDialog.getOpenFileName(
            bulk_email_window, "Select Excel File", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        selected_excel_path = file_name if file_name else None
        excel_path_label.setText(f"Selected Excel File: {file_name}" if file_name else "No Excel file selected.")

    def send_bulk_emails():
        if not selected_excel_path:
            QMessageBox.warning(bulk_email_window, "Error", "Please select an Excel file before proceeding.")
            return

        email_body_template = email_body_text.toPlainText().strip()
        if not email_body_template:
            QMessageBox.warning(bulk_email_window, "Error", "Email body cannot be empty.")
            return

        try:
            data = pd.read_excel(selected_excel_path)

            if data.empty:
                QMessageBox.warning(bulk_email_window, "Error", "The Excel file is empty.")
                return

            if "EMAIL" not in data.columns:
                QMessageBox.warning(bulk_email_window, "Error", "The Excel file must contain an 'EMAIL' column.")
                return

            outlook = win32.Dispatch("Outlook.Application")

            print("Sending emails...")
            for idx, row in data.iterrows():
                email = row.get("EMAIL", None)
                if not email or pd.isna(email):
                    print(f"Row {idx + 1}: No email address found, skipping.")
                    continue

                # Create a customized email body for the current row
                email_body = email_body_template
                for col_name in data.columns:
                    placeholder = f"[{col_name.upper()}]"
                    if placeholder in email_body:
                        col_value = str(row[col_name]) if not pd.isna(row[col_name]) else ""
                        email_body = email_body.replace(placeholder, col_value.upper() if col_name.upper() == "KEYWORD" else col_value)

                mail = outlook.CreateItem(0)
                mail.To = email
                mail.Subject = "Bulk Email Notification"
                mail.Body = email_body

                mail.Send()
                time.sleep(2)
                print(f"Row {idx + 1}: Email sent to {email}")

            QMessageBox.information(bulk_email_window, "Success", "All emails have been sent successfully.")

        except Exception as e:
            QMessageBox.critical(bulk_email_window, "Error", f"An error occurred: {str(e)}")

    def save_email_body():
    # Open a file dialog for the user to choose the save location
        file_name, _ = QFileDialog.getSaveFileName(
        bulk_email_window,
        "Save Email Body",
        "",
        "Text Files (*.txt);;All Files (*)"
    )
    
    # Check if the user canceled the save dialog
        if not file_name:
            return
        
        try:
            # Get the email body text from the QTextEdit widget
            email_body = email_body_text.toPlainText().strip()
            
            # Validate that the email body is not empty
            if not email_body:
                QMessageBox.warning(bulk_email_window, "Error", "Email body cannot be empty.")
                return
            
            # Write the email body to the chosen file
            with open(file_name, 'w', encoding='utf-8') as file:
                file.write(email_body)
            
            # Notify the user that the save was successful
            QMessageBox.information(bulk_email_window, "Success", f"Email body saved successfully to {file_name}")
        except Exception as e:
            # Handle any errors during the save process
            QMessageBox.critical(bulk_email_window, "Error", f"Failed to save email body: {str(e)}")

    def load_email_body():
        # Open a file dialog for the user to choose a file to load
        file_name, _ = QFileDialog.getOpenFileName(
            bulk_email_window,
            "Load Email Body",
            "",
            "Text Files (*.txt);;All Files (*)"
        )
        
        # Check if the user canceled the open dialog
        if not file_name:
            return

        try:
            # Read the content of the selected file
            with open(file_name, 'r', encoding='utf-8') as file:
                email_body = file.read()

            # Set the content of the QTextEdit widget
            email_body_text.setPlainText(email_body)

            # Notify the user that the load was successful
            QMessageBox.information(bulk_email_window, "Success", f"Email body loaded successfully from {file_name}")
        except Exception as e:
            # Handle any errors during the load process
            QMessageBox.critical(bulk_email_window, "Error", f"Failed to load email body: {str(e)}")

    app = QApplication.instance()
    if not app:  # Check if an instance of QApplication already exists
        app = QApplication([])

    bulk_email_window = QWidget()
    bulk_email_window.setWindowTitle("Bulk Email Sender")
    bulk_email_window.setMinimumSize(600, 500)
    bulk_email_window.setStyleSheet("background-color: #f4f4f4;")

    layout = QVBoxLayout()

    title_label = QLabel("Bulk Email System")
    title_label.setStyleSheet("font-size: 22px; font-weight: bold; color: #8b0000;")
    layout.addWidget(title_label, alignment=Qt.AlignCenter)

    excel_path_label = QLabel("No Excel file selected.")
    layout.addWidget(excel_path_label)

    excel_button = QPushButton("Select Excel File")
    excel_button.setStyleSheet("background-color: #8b0000; color: white; font-weight: bold; padding: 8px;")
    excel_button.clicked.connect(browse_excel)
    layout.addWidget(excel_button)

    # Add email body text box
    email_body_label = QLabel("Email Body:")
    email_body_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #333;")
    layout.addWidget(email_body_label)

    email_body_text = QTextEdit()
    email_body_text.setStyleSheet("""
        QTextEdit {
            background-color: #ffffff;
            border: 2px solid #8b0000;
            border-radius: 5px;
            font-size: 14px;
            color: #333;
            padding: 10px;
        }
        QTextEdit:focus {
            border: 2px solid #00aaff;
        }
    """)
    layout.addWidget(email_body_text)

    # Add "Save Email Body" and "Load Email Body" buttons
    email_buttons_layout = QHBoxLayout()

    save_button = QPushButton("Save Email Body")
    save_button.setStyleSheet("background-color: #008000; color: white; font-weight: bold; padding: 8px;")
    save_button.clicked.connect(save_email_body)
    email_buttons_layout.addWidget(save_button)

    load_button = QPushButton("Load Email Body")
    load_button.setStyleSheet("background-color: #1e90ff; color: white; font-weight: bold; padding: 8px;")
    load_button.clicked.connect(load_email_body)
    email_buttons_layout.addWidget(load_button)

    layout.addLayout(email_buttons_layout)

    send_button = QPushButton("Send Emails")
    send_button.setStyleSheet("background-color: #8b0000; color: white; font-weight: bold; padding: 10px;")
    send_button.clicked.connect(send_bulk_emails)
    layout.addWidget(send_button)

    bulk_email_window.setLayout(layout)
    bulk_email_window.show()
    app.exec_()


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
    title_label = QLabel("File Automation sSystem")
    title_label.setStyleSheet("font-size: 22px; font-weight: bold; color: #8b0000;")
    layout.addWidget(title_label, alignment=Qt.AlignCenter)

    # Labels for displaying selected file paths
    template_path_label = QLabel("No template selected.")
    excel_path_label = QLabel("No Excel file selected.")
    save_path_label = QLabel("No save directory selected.")
    for label in [template_path_label, excel_path_label, save_path_label]:
        label.setStyleSheet("font-size: 10px; color: #333;")
        layout.addWidget(label)

    # List to display saved automations with scrolling
    automation_list = QListWidget()
    automation_list.setStyleSheet("""
        font-size: 12px;
        background-color: #ffffff;
        border: 1px solid #ccc;
        padding: 5px;
    """)
    automation_list.setFixedHeight(100)
    automation_list.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
    layout.addWidget(automation_list)

    # Buttons for file selection and saving
    browse_template_button = QPushButton("Select Word Template")
    browse_excel_button = QPushButton("Select Excel File")
    browse_save_button = QPushButton("Select Save Directory")
    save_automation_button = QPushButton("Save Automation")
    delete_automation_button = QPushButton("Delete Selected Automation")
    run_automation_button = QPushButton("Run Automation")
    send_automation_button = QPushButton("Send Generated docs")

    button_style = """
        QPushButton {
            font-size: 14px;
            color: white;
            background-color: #8b0000;
            border: none;
            padding: 10px;
            margin: 3px 0;
        }
        QPushButton:hover {
            background-color: black;
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
    layout.addWidget(send_automation_button)

    # Variables to hold selected file paths

    # Function to browse and select a Word template
    def browse_template():
        global selected_template_path
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
        global selected_excel_path
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
        global save_directory_path
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
        global automations
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
    def send_saved_docs():
        print("Sending the files")
        global selected_excel_path  # Reference the global variable
        if not selected_excel_path:
            QMessageBox.warning(file_automation_window, "Error", "Please select an Excel file before proceeding.")
            return

        try:
            # Read the updated Excel file
            data = pd.read_excel(selected_excel_path)

            # Check if necessary columns exist
            required_columns = ["EMAIL", "Generated File Path", "NAME", "SURNAME"]
            for col in required_columns:
                if col not in data.columns:
                    QMessageBox.warning(
                        file_automation_window,
                        "Error",
                        f"The Excel file must contain the following columns: {', '.join(required_columns)}."
                    )
                    return

            outlook = win32.Dispatch("Outlook.Application")

            for idx, row in data.iterrows():
                email = row.get("EMAIL", None)
                file_path = row.get("Generated File Path", None)

                if not email or pd.isna(email) or not file_path or pd.isna(file_path):
                    print(f"Row {idx + 1}: Missing email or file path, skipping.")
                    continue

                # Create an email
                mail = outlook.CreateItem(0)
                mail.To = email
                mail.Subject = f"Generated Document for {row['NAME']} {row['SURNAME']}"
                mail.Body = "Please find the attached document."

                # Attach the file
                mail.Attachments.Add(file_path)

                mail.Send()
                print(f"Row {idx + 1}: Email sent to {email} with attachment {file_path}")

            QMessageBox.information(file_automation_window, "Success", "All documents have been sent successfully.")

        except Exception as e:
            QMessageBox.critical(file_automation_window, "Error", f"An error occurred: {str(e)}")

        
        
    # Connect buttons to functions
    browse_template_button.clicked.connect(browse_template)
    browse_excel_button.clicked.connect(browse_excel)
    browse_save_button.clicked.connect(browse_save_directory)
    save_automation_button.clicked.connect(save_automation)
    delete_automation_button.clicked.connect(delete_selected_automation)
    run_automation_button.clicked.connect(run_automation)
    send_automation_button.clicked.connect(send_saved_docs)

    # Load automations on startup
    load_automations()

    file_automation_window.setLayout(layout)
    file_automation_window.show()
    app.exec_()


if __name__ == "__main__":
    create_file_automation_window()
