# Import necessary libraries
from PyQt5.QtCore import Qt, QSize, QTimer
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QStackedWidget
)
from PyQt5.QtGui import QPixmap

import my_APC  # Import the test function
import Xero
import office_doc_automation
import Developer


# Initialization of the main application
def initialize_app():
    app = QApplication([])
    app.setStyleSheet(global_stylesheet())
    return app


# Stylesheet definitions
def global_stylesheet():
    return """
        QPushButton {
            padding: 10px;
            font-size: 16px;
        }
        QPushButton#sidebarButton {
            background-color: red;
            color: white;
            border: none;
            text-align: left;
            padding: 10px 20px;
        }
        QPushButton#sidebarButton:hover {
            background-color: darkred;
        }
        QLabel {
            font-size: 24px;
            margin: 20px;
        }
        QWidget#contentArea {
            background-color: white;
        }
        QPixmap{
        width: 200px
        height: 200px
        }
    """


# Splash Screen Implementation
class SplashScreen(QWidget):
    def __init__(self):
        super().__init__()
        self.setFixedSize(500, 500)
        self.setStyleSheet("background-color: #333; color: white;")
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        pixmap = QPixmap("Main\AHV1.0.3.png")
        splash_label = QLabel(self)
        splash_label.setPixmap(pixmap)
        splash_label.setAlignment(Qt.AlignCenter)

        title_label = QLabel("Automation Haven V1.0.3")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 28px; font-weight: bold; margin-top: 20px;")

        layout.addWidget(splash_label)
        layout.addWidget(title_label)
        self.setLayout(layout)


# Sidebar navigation
class SidebarNavigation(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        self.setFixedWidth(250)
        self.setStyleSheet("background-color: lightgray;")

        layout = QVBoxLayout()

        # Sidebar buttons
        buttons = [
            ("🏠 Home", "home_button"),
            ("📄 Xero Automation", "xero_button"),
            ("💼 Office Automation", "office_button"),
            ("📦 APC Package Tools", "apc_button"),
            ("⚙️ Settings", "settings_button"),
            ("⚙️ Developer", "Dev_Login_button"),
        ]
        for text, name in buttons:
            button = QPushButton(text, self)
            button.setObjectName("sidebarButton")
            setattr(self, name, button)  # Dynamically set attributes for each button
            layout.addWidget(button)

        layout.addStretch()
        self.setLayout(layout)


# Main window with sidebar and content
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Automation Haven V1.0.3")
         #self.setWindowFlags(Qt.FramelessWindowHint | Qt.Window)
        self.init_ui()

    def init_ui(self):
        # Main layout
        main_layout = QHBoxLayout(self)

        # Sidebar
        self.sidebar = SidebarNavigation(self)
        self.sidebar_visible = True

        # Sidebar toggle button
        self.toggle_button = QPushButton("☰", self)
        self.toggle_button.setFixedSize(QSize(50, 50))
        self.toggle_button.setStyleSheet("background-color: red; color: white; font-size: 18px; border: none;")
        self.toggle_button.clicked.connect(self.toggle_sidebar)

        # Content area
        self.content_area = QStackedWidget(self)
        self.content_area.setObjectName("contentArea")
        self.add_pages()

        # Connect sidebar buttons to pages
        self.sidebar.home_button.clicked.connect(lambda: self.content_area.setCurrentIndex(0))
        self.sidebar.xero_button.clicked.connect(lambda: self.content_area.setCurrentIndex(1))
        self.sidebar.office_button.clicked.connect(lambda: self.content_area.setCurrentIndex(2))
        self.sidebar.apc_button.clicked.connect(lambda: self.content_area.setCurrentIndex(3))
        self.sidebar.settings_button.clicked.connect(lambda: self.content_area.setCurrentIndex(4))
        self.sidebar.Dev_Login_button.clicked.connect(lambda: self.content_area.setCurrentIndex(5))

        # Layout setup
        main_layout.addWidget(self.sidebar)
        main_layout.addWidget(self.toggle_button)
        main_layout.addWidget(self.content_area)

    def toggle_sidebar(self):
        self.sidebar.setVisible(not self.sidebar_visible)
        self.sidebar_visible = not self.sidebar_visible

    def add_pages(self):
        # Add pages to the content area
        self.content_area.addWidget(self.create_page("Welcome to Automation Haven!"))
        self.content_area.addWidget(self.create_xero_page())
        self.content_area.addWidget(self.create_office_page())
        self.content_area.addWidget(self.create_apc_page())
        self.content_area.addWidget(self.create_page("Settings"))
        self.content_area.addWidget(self.create_Developer_page())

    def create_page(self, title, button_text=None):
        """Generic method for creating simple pages."""
        page = QWidget()
        layout = QVBoxLayout()
        label = QLabel(title)
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)

        if button_text:
            button = QPushButton(button_text)
            button.setFixedSize(200, 50)
            layout.addWidget(button, alignment=Qt.AlignCenter)

        page.setLayout(layout)
        return page

    def create_xero_page(self):
        page = QWidget()
        layout = QVBoxLayout()
        label = QLabel("Xero Automation")
        label.setAlignment(Qt.AlignCenter)

        xero_sender = QPushButton("Xero Statement Sender")
        xero_sender.setFixedSize(200, 50)
        xero_sender.clicked.connect(Xero.create_window)
        xero_Setup_Button = QPushButton("My Details")
        xero_Setup_Button.clicked.connect(Xero.xero_setup)

        layout.addWidget(label)
        layout.addWidget(xero_sender, alignment=Qt.AlignCenter)
        layout.addWidget(xero_Setup_Button, alignment=Qt.AlignCenter)
        page.setLayout(layout)
        return page

    def create_apc_page(self):
        page = QWidget()
        layout = QVBoxLayout()
        label = QLabel("APC Package Tools")
        label.setAlignment(Qt.AlignCenter)

        apc_automation = QPushButton("Keycodes and Pins")
        apc_script_automation = QPushButton("Script Sender")
        apc_automation.clicked.connect(my_APC.create_apc_window)
        apc_script_automation.clicked.connect(my_APC.create_script_window)

        layout.addWidget(label)
        layout.addWidget(apc_automation, alignment=Qt.AlignCenter)
        layout.addWidget(apc_script_automation, alignment=Qt.AlignCenter)
        page.setLayout(layout)
        return page

    def create_office_page(self):
        page = QWidget()
        layout = QVBoxLayout()
        label = QLabel("Office Automations")
        label.setAlignment(Qt.AlignCenter)

        office_file_automation_button = QPushButton("Doc Automations")
        bulk_email_button = QPushButton("Bulk Email Sender")
        bulk_email_button.clicked.connect(office_doc_automation.create_bulk_email_window)
        office_file_automation_button.clicked.connect(office_doc_automation.create_file_automation_window)

        layout.addWidget(label)
        layout.addWidget(office_file_automation_button, alignment=Qt.AlignCenter)
        layout.addWidget(bulk_email_button, alignment=Qt.AlignCenter)
        page.setLayout(layout)
        return page

    def create_Developer_page(self):
        page = QWidget()
        layout = QVBoxLayout()
        label = QLabel("Devs")
        label.setAlignment(Qt.AlignCenter)

        apc_automation = QPushButton("Log In")
        apc_automation.clicked.connect(Developer.create_login_window)

        layout.addWidget(label)
        layout.addWidget(apc_automation, alignment=Qt.AlignCenter)
        page.setLayout(layout)
        return page


# Main execution
if __name__ == "__main__":
    import sys

    app = initialize_app()

    # Show splash screen
    splash = SplashScreen()
    splash.show()

    # Timer to close splash and show the main window
    QTimer.singleShot(3000, splash.close)
    QTimer.singleShot(3000, lambda: MainWindow().showMaximized())

    sys.exit(app.exec_())
