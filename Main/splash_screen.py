# splash_screen.py
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QProgressBar
from PyQt5.QtGui import QPixmap, QFont
#This is a function to be included later on mabey after v1.0.06     

class SplashScreen(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Launching Automation Haven...")
        self.setFixedSize(500, 500)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setStyleSheet("background-color: #333; color: white; border-radius: 10px;")
        self.init_ui()

    def init_ui(self):
        # Main layout
        layout = QVBoxLayout()

        # Image
        pixmap = QPixmap("Main/AHV1.0.3.png")  # Replace with the actual path to your image
        splash_label = QLabel(self)
        splash_label.setPixmap(pixmap)
        splash_label.setAlignment(Qt.AlignCenter)

        # Title
        title_label = QLabel("Automation Haven V1.0.3")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; margin-top: 20px;")

        # Subtitle
        subtitle_label = QLabel("Initializing, please wait...")
        subtitle_label.setAlignment(Qt.AlignCenter)
        subtitle_label.setFont(QFont("Arial", 12))
        subtitle_label.setStyleSheet("font-size: 14px; margin-top: 10px;")

        # Progress Bar
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                background-color: #555;
                border-radius: 5px;
                height: 10px;
            }
            QProgressBar::chunk {
                background-color: orange;
                border-radius: 5px;
            }
        """)

        # Add widgets to layout
        layout.addWidget(splash_label)
        layout.addWidget(title_label)
        layout.addWidget(subtitle_label)
        layout.addWidget(self.progress_bar)
        layout.setSpacing(10)

        self.setLayout(layout)

    def update_progress(self, value):
        """Update the progress bar."""
        self.progress_bar.setValue(value)
"""
if __name__ == "__main__":
    import sys

    app = QApplication([])

    # Show the splash screen
    splash = splash_screen.SplashScreen()
    splash.show()

    # Simulate loading progress
    for i in range(1, 101):
        QTimer.singleShot(i * 30, lambda val=i: splash.update_progress(val))

    # Close the splash screen and show the main window
    QTimer.singleShot(3000, splash.close)  # Close the splash screen after 3 seconds
    QTimer.singleShot(3000, lambda: initialize_app().showMaximized())  # Show the main window after 3 seconds

    sys.exit(app.exec_())

"""