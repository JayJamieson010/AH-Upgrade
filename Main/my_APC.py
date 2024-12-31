from PyQt5.QtWidgets import (
    QWidget, QLabel, QVBoxLayout, QPushButton, QFileDialog, QMessageBox, QComboBox
)
import pandas as pd
import win32com.client as win32
import time
import os


# Global variable to keep the APC window reference
apc_window = None

#Add the create_Script_window here
def create_script_window():
    global apc_window
    selected_folder_path = None
    selected_excel_path = None

    def browse_folder():
        nonlocal selected_folder_path
        folder_path = QFileDialog.getExistingDirectory(apc_window, "Select Folder")
        selected_folder_path = folder_path if folder_path else None
        folder_path_label.setText(f"Selected Folder: {folder_path}" if folder_path else "No folder selected.")

    def browse_excel():
        nonlocal selected_excel_path
        file_name, _ = QFileDialog.getOpenFileName(
            apc_window, "Select Excel File", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        selected_excel_path = file_name if file_name else None
        excel_path_label.setText(f"Selected Excel File: {file_name}" if file_name else "No Excel file selected.")

    def run_sender():
    
        if not selected_excel_path or not selected_folder_path:
            QMessageBox.warning(apc_window, "Error", "Please select both an Excel file and a folder before running.")
            return

        try:
            # Read the Excel file
            data = pd.read_excel(selected_excel_path)

            if data.empty:
                QMessageBox.warning(apc_window, "Error", "The Excel file is empty.")
                return

            # Ensure the STUDENTNUMBER column exists
            if 'STUDENTNUMBER' not in data.columns:
                QMessageBox.warning(apc_window, "Error", "The Excel file must contain a 'STUDENTNUMBER' column.")
                return

            # Add columns to store paths for the different file types
            if 'PDF_PATH' not in data.columns:
                data['PDF_PATH'] = None
            if 'DOCX_PATH' not in data.columns:
                data['DOCX_PATH'] = None
            if 'EXCEL_PATH' not in data.columns:
                data['EXCEL_PATH'] = None

            # Process each row and look for the corresponding files
            print("Processing all rows in the Excel file:")
            for idx, row in data.iterrows():
                student_number = row['STUDENTNUMBER']

                if pd.isna(student_number):
                    print(f"Row {idx + 1}: Missing STUDENTNUMBER, skipping.")
                    continue

                # Convert the STUDENTNUMBER to a string to handle both numeric and alphanumeric cases
                student_number_str = str(student_number).strip()

                # Construct file names for each type
                pdf_file_name = f"{student_number_str}.pdf"
                docx_file_name = f"{student_number_str}.docx"
                excel_file_name = f"{student_number_str}.xlsx"

                # Paths for each file type
                pdf_file_path = os.path.join(selected_folder_path, pdf_file_name)
                docx_file_path = os.path.join(selected_folder_path, docx_file_name)
                excel_file_path = os.path.join(selected_folder_path, excel_file_name)

                # Check if the files exist and update the DataFrame
                if os.path.exists(pdf_file_path):
                    print(f"Row {idx + 1}: Found PDF file for STUDENTNUMBER {student_number_str}: {pdf_file_path}")
                    data.at[idx, 'PDF_PATH'] = pdf_file_path
                else:
                    print(f"Row {idx + 1}: PDF file for STUDENTNUMBER {student_number_str} not found.")

                if os.path.exists(docx_file_path):
                    print(f"Row {idx + 1}: Found DOCX file for STUDENTNUMBER {student_number_str}: {docx_file_path}")
                    data.at[idx, 'DOCX_PATH'] = docx_file_path
                else:
                    print(f"Row {idx + 1}: DOCX file for STUDENTNUMBER {student_number_str} not found.")

                if os.path.exists(excel_file_path):
                    print(f"Row {idx + 1}: Found Excel file for STUDENTNUMBER {student_number_str}: {excel_file_path}")
                    data.at[idx, 'EXCEL_PATH'] = excel_file_path
                else:
                    print(f"Row {idx + 1}: Excel file for STUDENTNUMBER {student_number_str} not found.")

            # Save the updated Excel file
            updated_excel_path = os.path.join(os.path.dirname(selected_excel_path), "Updated_" + os.path.basename(selected_excel_path))
            data.to_excel(updated_excel_path, index=False)
            QMessageBox.information(apc_window, "Success", f"Processing complete. Updated file saved as:\n{updated_excel_path}")
            send_email(updated_excel_path)
           
            
        except Exception as e:
            QMessageBox.critical(apc_window, "Error", f"An error occurred while processing the file: {str(e)}")

     # Process Emails

    def send_email(updated_excel_path):
        try:
            # Load the updated Excel file into a DataFrame
            data = pd.read_excel(updated_excel_path)

            # Initialize Outlook
            outlook = win32.Dispatch("Outlook.Application")

            print("Starting email sending process...")

            # Loop through each row in the DataFrame
            for idx, row in data.iterrows():
                email = row.get("EMAIL", None)
                if not email or pd.isna(email):
                    print(f"Row {idx + 1}: No email address found, skipping.")
                    continue

                # Create a new email
                mail = outlook.CreateItem(0)  # 0 corresponds to MailItem
                mail.To = email
                mail.Subject = "Your Apc Script"
                mail.Body = f"Dear {row.get('STUDENTNUMBER', 'Student')},\n\nPlease find the attached files.\n\nBest regards,\nJay"

                files_attached = False  # Track whether any files were attached

                # Attach files based on available paths
                for file_type, column_name in [("PDF", "PDF_PATH"), ("DOCX", "DOCX_PATH"), ("EXCEL", "EXCEL_PATH")]:
                    file_path = row.get(column_name, None)

                    # Ensure the file path is valid and check if it exists
                    if isinstance(file_path, float) and pd.isna(file_path):  # Handle NaN values
                        file_path = None
                    elif file_path:  # Convert to string if not NaN
                        file_path = str(file_path).strip()

                    if file_path and os.path.exists(file_path):
                        mail.Attachments.Add(file_path)
                        print(f"Row {idx + 1}: Attached {file_type} file {file_path}")
                        files_attached = True
                    else:
                        print(f"Row {idx + 1}: {file_type} file not found or path invalid.")

                # If no files were attached, skip sending the email
                if not files_attached:
                    print(f"Row {idx + 1}: No valid files to attach. Email skipped for {email}.")
                    continue

                # Send the email
                mail.Send()
                print(f"Row {idx + 1}: Email sent to {email}")

            QMessageBox.information(apc_window, "Success", "All emails have been sent successfully.")

        except Exception as e:
            QMessageBox.critical(apc_window, "Error", f"An error occurred while sending emails: {str(e)}")



    apc_window = QWidget()
    apc_window.setWindowTitle("APC Automation Tool")
    apc_window.setMinimumSize(400, 400)

    # Layout and UI
    layout = QVBoxLayout()
    layout.addWidget(QLabel("APC Script Sender Automation Tool").setStyleSheet("font-size: 18px; font-weight: bold;"))
    layout.addWidget(QLabel("Required Excel Fields:\n- STUDENTNUMBER\n- EMAIL").setStyleSheet("font-size: 14px;"))

    # Folder/Excel Selection
    excel_path_label = QLabel("No Excel file selected.")
    folder_path_label = QLabel("No folder selected.")
    layout.addWidget(excel_path_label)
    layout.addWidget(folder_path_label)

    excel_button = QPushButton("Select Excel File")
    folder_button = QPushButton("Select Folder")
    excel_button.clicked.connect(browse_excel)
    folder_button.clicked.connect(browse_folder)
    layout.addWidget(excel_button)
    layout.addWidget(folder_button)

    # Run Button
    run_button = QPushButton("Run Script")
    run_button.clicked.connect(run_sender)
    layout.addWidget(run_button)

    # Finalize Window
    apc_window.setLayout(layout)
    apc_window.show()




def create_apc_window():
    """Function to create and display the APC Automation window."""
    global apc_window  # Use a global variable to hold the window reference

    # Create a new widget as a window
    apc_window = QWidget()
    apc_window.setWindowTitle("APC Script Automation")
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
