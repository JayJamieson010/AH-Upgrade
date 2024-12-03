from PyQt5.QtWidgets import (
    QWidget, QLabel, QVBoxLayout, QPushButton, QFileDialog, QMessageBox
)
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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
    def clear_and_write(content):
        try:
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.5)
            pyautogui.press('backspace')
            time.sleep(0.5)

            print(f"Writing content: {content}")
            pyautogui.write(content)
            time.sleep(1)
            print(f"Successfully wrote content: {content}")
        except Exception as e:
            print(f"Error writing content: {e}")

    def locate_element(image_path, confidence=0.8, timeout=10):
        print(f"Locating element: {image_path}")
        start_time = time.time()
        while time.time() - start_time < timeout:
            element = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
            if element:
                print(f"Element {image_path} located.")
                return element
            time.sleep(0.5)
        print(f"Failed to locate element: {image_path}")
        return None


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

            driver.execute_script("window.scrollBy(0, 500);")
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "xui-margin-bottom-xsmall")))

            # Check for MFA requirement
            mfa_element = driver.find_elements(By.CLASS_NAME, "xui-margin-bottom-xsmall")
            if mfa_element:
                pyautogui.scroll(-200)
                use_backup = pyautogui.locateCenterOnScreen(r"Main\Images\Use_Backup.png", confidence=0.7)
                if use_backup:
                    pyautogui.click(use_backup)
                    backup_email = pyautogui.locateCenterOnScreen(r"Main\Images\Backup email.png", confidence=0.7)
                    pyautogui.click(backup_email)
                    send_backup_code = pyautogui.locateCenterOnScreen(r"Main\Images\send_backup_code.png", confidence=0.7)
                    pyautogui.click(send_backup_code)
                    time.sleep(30)
                    pyautogui.press("tab")
                    pyautogui.press("enter")

                time.sleep(20)# Wait for page to load 
                contacts = locate_element("Main\Images\Contacts.png", confidence=0.8)
                if contacts:
                         pyautogui.click(contacts)
                         time.sleep(2)
                        

                         all_contacts = locate_element("Main\Images\AllContacts.png", confidence=0.8)

                if all_contacts:
                         pyautogui.click(all_contacts)
                         time.sleep(3)

                         options = locate_element("Main\Images\options.png", confidence=0.8)
                if options:
                         pyautogui.click(options)
                         time.sleep(3)

                         send_statement = locate_element("Main\Images\sendStatement.png", confidence=0.8)
                if send_statement:
                         pyautogui.click(send_statement)
                         time.sleep(3)

                    # Additional Xero steps with Excel data integration
            

            # Further automation process...
            # Loop through the student list and perform necessary actions
            for student_tuple in student_list:
                student_number = student_tuple[0]
                if student_number is None:
                    continue
                print(f"Processing student number: {student_number}")
                time.sleep(4)

                activity_field = driver.find_element(By.ID, "StatementTypeFromForm_value")
                if activity_field:
                            print("Found")
                            activity_field.click()
                            # Ctrl + A and Backspace using Selenium
                            activity_field.send_keys(Keys.CONTROL + 'a')  # Select all text
                            activity_field.send_keys(Keys.BACKSPACE)  # Clear the text
                            activity_field.send_keys("Activity")  # Type "Activity
                                                

                            # Navigate through fields and input data
                            pyautogui.press('tab')
                            pyautogui.write("1 Jan 2023")
                            pyautogui.press('tab')
                            pyautogui.write(end_date_selected)
                            pyautogui.press('tab')
                            clear_and_write(str(student_number))
                            pyautogui.press("tab")
                            pyautogui.press("enter")

                            time.sleep(3)

                            checkboxes = driver.find_elements(By.ID, "ext-gen26")

                            # Loop through each checkbox and click it if it is not already selected
                            for checkbox in checkboxes:
                                if not checkbox.is_selected():
                                    checkbox.click()

                            email = driver.find_element(By.ID, "ext-gen22")

                            if email:
                                email.click()
                                time.sleep(3)
                            
                            emailAdress = driver.find_element(By.XPATH, "//*[contains(@id, 'MessageTo')]")

                            if emailAdress:
                                emailAdress.click()
                                emailAdress.send_keys(Keys.CONTROL + 'a')  # Select all text
                                emailAdress.send_keys(Keys.BACKSPACE)
                                emailAdress.send_keys("Jay@ias.ac.za")

                            sendButton= driver.find_element(By.ID, "email01")
                            if sendButton:
                                sendButton.click()
                                
                            else:
                               time.sleep(20)# Wait for page to load 
            else:
                time.sleep(20)# Wait for page to load 
                contacts = locate_element("Main\Images\Contacts.png", confidence=0.8)
                if contacts:
                         pyautogui.click(contacts)
                         time.sleep(2)
                        

                         all_contacts = locate_element("Main\Images\AllContacts.png", confidence=0.8)

                if all_contacts:
                         pyautogui.click(all_contacts)
                         time.sleep(3)

                         options = locate_element("Main\Images\options.png", confidence=0.8)
                if options:
                         pyautogui.click(options)
                         time.sleep(3)

                         send_statement = locate_element("Main\Images\sendStatement.png", confidence=0.8)
                if send_statement:
                         pyautogui.click(send_statement)
                         time.sleep(3)

                    # Additional Xero steps with Excel data integration
            

            # Further automation process...
            # Loop through the student list and perform necessary actions
            for student_tuple in student_list:
                student_number = student_tuple[0]
                if student_number is None:
                    continue
                print(f"Processing student number: {student_number}")
                time.sleep(4)

                activity_field = driver.find_element(By.ID, "StatementTypeFromForm_value")
                if activity_field:
                            print("Found")
                            activity_field.click()
                            # Ctrl + A and Backspace using Selenium
                            activity_field.send_keys(Keys.CONTROL + 'a')  # Select all text
                            activity_field.send_keys(Keys.BACKSPACE)  # Clear the text
                            activity_field.send_keys("Activity")  # Type "Activity
                                                

                            # Navigate through fields and input data
                            pyautogui.press('tab')
                            pyautogui.write("1 Jan 2023")
                            pyautogui.press('tab')
                            pyautogui.write(end_date_selected)
                            pyautogui.press('tab')
                            clear_and_write(str(student_number))
                            pyautogui.press("tab")
                            pyautogui.press("enter")

                            time.sleep(3)

                            checkboxes = driver.find_elements(By.ID, "ext-gen26")

                            # Loop through each checkbox and click it if it is not already selected
                            for checkbox in checkboxes:
                                if not checkbox.is_selected():
                                    checkbox.click()

                            email = driver.find_element(By.ID, "ext-gen22")

                            if email:
                                email.click()
                                time.sleep(3)
                            
                            emailAdress = driver.find_element(By.XPATH, "//*[contains(@id, 'MessageTo')]")

                            if emailAdress:
                                emailAdress.click()
                                emailAdress.send_keys(Keys.CONTROL + 'a')  # Select all text
                                emailAdress.send_keys(Keys.BACKSPACE)
                                emailAdress.send_keys("Jay@ias.ac.za")

                            sendButton= driver.find_element(By.ID, "email01")
                            if sendButton:
                                sendButton.click()
                                
                            else:
                               time.sleep(20)# Wait for page to load 
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
