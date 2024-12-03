from PyQt5.QtWidgets import (
    QWidget, QLabel, QVBoxLayout, QPushButton, QFileDialog, QMessageBox, QComboBox
)
import pandas as pd
import win32com.client as win32
import time

# Global variable to keep the APC window reference
apc_window = None

def create_apc_window():
    """Function to create and display the APC Automation window."""
    global apc_window  # Use a global variable to hold the window reference

    # Create a new widget as a window
    apc_window = QWidget()
    apc_window.setWindowTitle("APC Automation Window")
    apc_window.setMinimumSize(400, 400)

    # Variable to hold the selected file path and pin_only_keycode
    selected_file_path = None
    pin_only_keycode = "keycode"  # Default value

    # Function to browse and select an Excel file
    def browse_file():
        nonlocal selected_file_path
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(
            apc_window, "Select Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options
        )
        if file_name:
            excel_path_label.setText(f"Selected File: {file_name}")
            selected_file_path = file_name
        else:
            excel_path_label.setText("No file selected.")
            selected_file_path = None

    # Function to update the pin_only_keycode value based on dropdown selection
    def update_pin_only_keycode(value):
        nonlocal pin_only_keycode
        if value == "Keycode Only":
            pin_only_keycode = "keycode"
        else:
            pin_only_keycode = ""

    # Function to send APC codes using the selected Excel file
    def send_apc_codes_gui():
        """Runs the APC code automation process."""
        if not selected_file_path:
            QMessageBox.warning(apc_window, "Warning", "Please select an Excel file.")
            return

        try:
            send_apc_codes(pin_only_keycode=pin_only_keycode, apc_excel_path=selected_file_path)
            QMessageBox.information(apc_window, "Success", "APC automation completed successfully.")
        except Exception as e:
            print(f"An error occurred: {e}")
            QMessageBox.critical(apc_window, "Error", f"An error occurred during the automation: {e}")

    # Create the layout and add components
    layout = QVBoxLayout()

    # Title label
    title_label = QLabel("APC Automation Tool")
    title_label.setStyleSheet("font-size: 18px; font-weight: bold;")
    layout.addWidget(title_label)

    # Description of required fields
    description_label = QLabel(
        "Required Excel Fields:\n"
        "- Student Number\n"
        "- Name\n"
        "- Surname\n"
        "- Email\n"
        "- Keycode\n"
        "- Pins (optional, depending on selected mode)\n"
        "- Subject\n"
        "- Processed\n"
    )
    description_label.setStyleSheet("font-size: 14px; margin-bottom: 10px;")
    layout.addWidget(description_label)

    # Label to display the selected file
    excel_path_label = QLabel("No file selected.")
    excel_path_label.setStyleSheet("font-size: 16px;")
    layout.addWidget(excel_path_label)

    # Dropdown menu for mode selection
    mode_dropdown = QComboBox()
    mode_dropdown.addItems(["Keycode Only", "Keycodes and Pins"])
    mode_dropdown.setStyleSheet("font-size: 14px; margin-bottom: 10px;")
    mode_dropdown.currentTextChanged.connect(update_pin_only_keycode)
    layout.addWidget(QLabel("Select Mode:").setStyleSheet("font-size: 14px;"))
    layout.addWidget(mode_dropdown)

    # Buttons for browsing and running automation
    browse_button = QPushButton("Browse Excel File")
    browse_button.clicked.connect(browse_file)
    browse_button.setStyleSheet("font-size: 14px; margin-bottom: 10px;")
    layout.addWidget(browse_button)

    run_sender_button = QPushButton("Run APC Automation")
    run_sender_button.clicked.connect(send_apc_codes_gui)
    run_sender_button.setStyleSheet("font-size: 14px;")
    layout.addWidget(run_sender_button)

    # Set the layout and show the window
    apc_window.setLayout(layout)
    apc_window.show()

# The APC automation function
def send_apc_codes(pin_only_keycode="", apc_excel_path=None):
    if not apc_excel_path:
        apc_excel_path = r"C:\Users\jayja\OneDrive\Documents\AH Files\APC AH.xlsx"

    # Load Excel file
    try:
        df = pd.read_excel(apc_excel_path)
    except FileNotFoundError:
        print(f"File not found: {apc_excel_path}")
        return
    except Exception as e:
        print(f"Error loading file: {e}")
        return

    # Clean column names
    df.columns = df.columns.str.strip()

    # Validate required columns
    required_columns = ["Student Number", "Name", "Surname", "Email", "Keycode", "Body", "Notice", "Subject", "Processed"]
    for col in required_columns:
        if col not in df.columns:
            print(f"Missing column: {col}")
            return

    # Initialize Outlook
    try:
        outlook = win32.Dispatch('outlook.application')
    except Exception as e:
        print(f"Failed to initialize Outlook: {e}")
        return

    # Send emails
    for _, row in df.iterrows():
        student_number = str(row["Student Number"])
        name = row["Name"]
        email = row["Email"].strip()
        keycode = row["Keycode"]
        pin = row.get("Pin", "N/A")
        subject = row["Subject"]

        if pin_only_keycode == "keycode":
            body = (
                f"Dear {name},\n\n"
                f"Your student number is {student_number}.\n\n"
                f"Keycode: {keycode}\n\n"
                "Best regards,\n\nYour Team"
            )
        else:
            body = (
                f"Dear {name},\n\n"
                f"Your student number is {student_number}.\n\n"
                f"Keycode: {keycode}\n"
                f"Pins: {pin}\n\n"
                "Best regards,\n\nYour Team"
            )

        try:
            mail = outlook.CreateItem(0)
            mail.To = email
            mail.Subject = subject
            mail.Body = body
            mail.Send()
            print(f"Email sent to {email}")
            time.sleep(2)
        except Exception as e:
            print(f"Failed to send email to {email}: {e}")
