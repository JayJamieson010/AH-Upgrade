from PyQt5.QtWidgets import (
    QWidget, QLabel, QVBoxLayout, QPushButton, QFileDialog, QMessageBox
)
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook
import pyautogui
import time


# Global variable to keep the window reference
xero_window = None

def create_window():
    """Function to create and display the Xero Automation window."""
    global xero_window  # Use a global variable to hold the window reference

    # Create a new widget as a window
    xero_window = QWidget()
    xero_window.setWindowTitle("Xero Automation Window")
    xero_window.setMinimumSize(400, 300)

    # Variables to hold selected file path
    selected_file_path = None

    # Function to browse and select an Excel file
    def browse_file():
        nonlocal selected_file_path
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(
            xero_window, "Select Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options
        )
        if file_name:
            excel_path_label.setText(f"Selected File: {file_name}")
            selected_file_path = file_name
        else:
            excel_path_label.setText("No file selected.")
            selected_file_path = None

    # Xero Automation function
    def xero_statement_sender():
        """Runs the Xero automation process."""
        if not selected_file_path:
            QMessageBox.warning(xero_window, "Warning", "Please select an Excel file.")
            return

        try:
            # Path to your chromedriver.exe file
            PATH = r"C:\webdrivers\chromedriver-win64\chromedriver.exe"
            service = Service(PATH)
            driver = webdriver.Chrome(service=service)

            wb = load_workbook(selected_file_path)
            ws = wb.active
            student_list = list(ws.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True))
            end_date_selected = pyautogui.prompt("End Date?", "Text")
            time.sleep(3)

            # Maximize the browser window
            driver.maximize_window()

            # Navigate to the Xero login page
            driver.get("https://login.xero.com/")
            time.sleep(3)

            # Login process
            email_field = driver.find_element(By.CLASS_NAME, "xui-textinput--input")
            password_field = driver.find_element(By.CLASS_NAME, "xl-form-password")
            email_field.send_keys("JayJamieson010@gmail.com")
            password_field.send_keys("KanoDagurHeather4fam")
            login_button = driver.find_element(By.CLASS_NAME, "xui-button")
            login_button.click()
            time.sleep(5)

            # Further automation process...
            # Loop through the student list and perform necessary actions
            for student_tuple in student_list:
                student_number = student_tuple[0]
                if student_number is None:
                    continue
                print(f"Processing student number: {student_number}")
                # Add the rest of your automation logic here...

            QMessageBox.information(xero_window, "Success", "Xero automation completed successfully.")
            driver.quit()

        except Exception as e:
            print(f"An error occurred: {e}")
            QMessageBox.critical(xero_window, "Error", f"An error occurred during the automation: {e}")

    # Create the layout and add components
    layout = QVBoxLayout()

    # Label to display the selected file
    excel_path_label = QLabel("No file selected.")
    excel_path_label.setStyleSheet("font-size: 16px;")

    # Buttons for browsing and running automation
    browse_button = QPushButton("Browse Excel File")
    browse_button.clicked.connect(browse_file)

    run_sender_button = QPushButton("Run Automation")
    run_sender_button.clicked.connect(xero_statement_sender)

    # Add components to the layout
    layout.addWidget(QLabel("This is the Xero Automation window.").setStyleSheet("font-size: 18px;"))
    layout.addWidget(excel_path_label)
    layout.addWidget(browse_button)
    layout.addWidget(run_sender_button)

    # Set the layout and show the window
    xero_window.setLayout(layout)
    xero_window.show()
