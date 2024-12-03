from PyQt5.QtWidgets import (
    QWidget, QLabel, QVBoxLayout, QLineEdit, QPushButton, QMessageBox, QApplication
)
import live_code_editor  # Import the Live Code Editor

# Global variables to keep window references
login_window = None
developer_page = None  # Reference for the Developer Page window


def create_login_window():
    """Function to create and display the Login Window."""
    global login_window  # Use a global variable to hold the window reference

    # Create a new widget as the login window
    login_window = QWidget()
    login_window.setWindowTitle("Login Page")
    login_window.setMinimumSize(400, 300)

    # Variables for storing user input
    user_email = ""
    user_password = ""

    # Function to handle the login process
    def handle_login():
        nonlocal user_email, user_password

        # Retrieve user input from the fields
        user_email = email_input.text()
        user_password = password_input.text()

        if not user_email or not user_password:
            QMessageBox.warning(login_window, "Input Error", "Please enter both email and password.")
            return

        # Example authentication logic (replace with your actual logic)
        if user_email == "admin@example.com" and user_password == "password123":
            QMessageBox.information(login_window, "Login Successful", "Welcome to the app!")
            login_window.close()  # Close the login window after successful login
            create_developer_page()  # Open the Developer Page
        else:
            QMessageBox.critical(login_window, "Login Failed", "Invalid email or password. Please try again.")

    # Create the layout and add components
    layout = QVBoxLayout()

    # Login label
    login_label = QLabel("Login to Your Account")
    login_label.setStyleSheet("font-size: 18px; font-weight: bold; text-align: center;")
    layout.addWidget(login_label)

    # Email input field
    email_input = QLineEdit()
    email_input.setPlaceholderText("Enter your email address")
    email_input.setStyleSheet("font-size: 14px;")
    layout.addWidget(email_input)

    # Password input field
    password_input = QLineEdit()
    password_input.setPlaceholderText("Enter your password")
    password_input.setEchoMode(QLineEdit.Password)  # Mask the input for passwords
    password_input.setStyleSheet("font-size: 14px;")
    layout.addWidget(password_input)

    # Login button
    login_button = QPushButton("Login")
    login_button.setStyleSheet("font-size: 16px;")
    login_button.clicked.connect(handle_login)
    layout.addWidget(login_button)

    # Set the layout and show the window
    login_window.setLayout(layout)
    login_window.show()


def create_developer_page():
    """Function to create and display the Developer Page."""
    global developer_page  # Use a global variable to keep the developer page reference

    # Create a new widget as the Developer Page
    developer_page = QWidget()
    developer_page.setWindowTitle("Developer Page")
    developer_page.setMinimumSize(400, 300)

    # Function to expand the window and add buttons
    def expand_window_and_add_buttons():
        # Increase the window size
        developer_page.resize(600, 500)

        # Add advanced feature buttons dynamically
        advanced_button_1 = QPushButton("Live Code Editor")
        advanced_button_1.setStyleSheet("font-size: 16px;")
        advanced_button_1.clicked.connect(live_code_editor.open_editor)  # Attach the Live Code Editor function
        layout.addWidget(advanced_button_1)

        advanced_button_2 = QPushButton("Set Up Global Variables")
        advanced_button_2.setStyleSheet("font-size: 16px;")
        layout.addWidget(advanced_button_2)

        # Disable the expand button to prevent further clicks
        expand_button.setDisabled(True)

        QMessageBox.information(developer_page, "Advanced Features Unlocked", "Advanced features are now available!")

    # Function to open the Live Code Editor
    def open_live_code_editor():
        """Function to open the Live Code Editor."""
        code_editor = live_code_editor.open_editor()  # Create an instance of the Live Code Editor
        code_editor.show()  # Show the Live Code Editor window

    # Layout for the Developer Page
    layout = QVBoxLayout()

    # Welcome label
    welcome_label = QLabel("Welcome to the Developer Page")
    welcome_label.setStyleSheet("font-size: 18px; font-weight: bold; text-align: center;")
    layout.addWidget(welcome_label)

    # Button to expand window and add features
    expand_button = QPushButton("Expand for Advanced Features")
    expand_button.setStyleSheet("font-size: 16px;")
    expand_button.clicked.connect(expand_window_and_add_buttons)
    layout.addWidget(expand_button)

    # Set the layout and show the window
    developer_page.setLayout(layout)
    developer_page.show()


# Entry point to run the app
if __name__ == "__main__":
    import sys

    app = QApplication(sys.argv)

    # Show the Login Window
    create_login_window()

    sys.exit(app.exec_())
