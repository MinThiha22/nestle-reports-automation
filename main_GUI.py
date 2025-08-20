from datetime import datetime
import sys
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QFrame
)

from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QPropertyAnimation, QEasingCurve
from PyQt6.QtGui import QFont, QPalette, QIcon
import os

from playwright.sync_api import sync_playwright
from dotenv import load_dotenv
import time

from excel import excel_process
from manipulateExcel import extract_zip
from checkDownload import run_check
from exportFile import export_file 
from importFile import import_file
from login import login_humby 

FIRST_INTERVAL_SECONDS = 3 * 60 * 60  # 3 hour wait to check for file download completion
SECOND_INTERVAL_SECONDS = 0.5 * 60 * 60  # 30 minutes wait to check for file download completion

# ====== Download folder and user data directory setup ======
downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
os.makedirs(downloads_folder, exist_ok=True)
user_data_dir = os.path.join(os.getcwd(), "user-data")

# === Step.1 Threaded worker to avoid freezing UI ===
class DownloadWorker(QThread):
    finished = pyqtSignal()
    log = pyqtSignal(str)

    def run(self):
        try:
            self.log.emit("🔐 Launching browser and logging in...")
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=False, args=["--start-maximized"])
                context = browser.new_context(no_viewport=True, accept_downloads=True)
                page = login_humby(context)
                self.log.emit("✅ Login successful, starting export...")
                export_file(page)
                self.log.emit("✅ Export complete.")
                context.close()
                browser.close()
        except Exception as e:
            self.log.emit(f"❌ Error during download: {e}")
        finally:
            self.finished.emit()

# === Step.2 Threaded worker to avoid freezing UI ===
class ImportAndDownloadFile(QThread):
    finished = pyqtSignal()
    log = pyqtSignal(str)
    error = pyqtSignal(str)
    import_success = pyqtSignal(str)
    import_failed = pyqtSignal(str, str)

    def __init__(self, file_path, parent=None):
        super().__init__(parent)
        self.file_path = file_path

    def run(self):
        try:
            self.log.emit("🔐 Launching browser and logging in...")
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=False, args=["--start-maximized"])
                context = browser.new_context(no_viewport=True, accept_downloads=True)
                page = login_humby(context)
                self.log.emit("✅ Login successful, starting Importing File...")
                
                result = import_file(page, file_path=self.file_path)
                if result.success:
                    self.log.emit("✅ Import completed successfully.")
                    self.import_success.emit(result.message)
                    self.log.emit("✅ Resubmit process done, 🔁 Wait for status complete..")
                else:
                    self.log.emit(f"❌ Import failed: {result.message}")
                    self.error.emit(result.message)
                    self.import_failed.emit(result.message, result.has_error)

                context.close()
                browser.close()
        except Exception as e:
            self.log.emit(f"❌ Error during download: {e}")
        finally:
            self.finished.emit()

# === Step.3 Threaded worker to avoid freezing UI ===
class DownloadProductFile(QThread):
    log = pyqtSignal(str)
    finished = pyqtSignal()

    def run(self):
        while True:
            try:
                self.log.emit("🔐 Launching browser and logging in...")
                with sync_playwright() as p:
                    context = p.chromium.launch_persistent_context(
                        user_data_dir=user_data_dir,
                        headless=False,
                        args=["--start-maximized"],
                        accept_downloads=True,
                        downloads_path=downloads_folder,
                    )
                    browser = context.browser
                    downloaded_path = self.run_download_updated_file(context)
                    context.close()
                    browser.close()
                    if downloaded_path:
                        self.process_download(downloaded_path)
                        return  # ✅ Exit the loop
            except Exception as e:
                self.log.emit(f"❌ Error during Playwright session: {e}")
    
    # This method unpacks the downloaded file if it is a zip file        
    def process_download(self, downloaded_path):
        """Simple version using the basic extract_zip_simple function"""
        if not downloaded_path:
            self.log.emit("❌ No file was downloaded.")
            return
            
        if downloaded_path.endswith(".zip"):
            self.log.emit(f"📁 Downloaded zip file at: {downloaded_path}")
            extract_folder = os.path.splitext(downloaded_path)[0]
            self.log.emit(f"🔄 Extracting to: {extract_folder}")
            result = extract_zip(downloaded_path, extract_folder)
            if result.success:
                self.log.emit(f"✅ Extraction completed successfully")
                self.log.emit(result.message)
                self.log.emit("🎉 Download completed successfully!")
                # Close the application after 3 seconds
                self.log.emit("🔄 Closing application in 3 seconds...")
                QTimer.singleShot(3000, self.close_application)
            else:
                error_msg = f"❌ Failed to extract: {downloaded_path}"
                self.log.emit(error_msg)
                self.error.emit(error_msg)
        else:
            self.log.emit(f"📁 File downloaded: {downloaded_path}")
            self.log.emit("🎉 Download completed successfully!")
            # Close the application after 3 seconds
            self.log.emit("🔄 Closing application in 3 seconds...")
            QTimer.singleShot(3000, self.close_application)
    
    # This method closes the application after the download is complete
    # and the extraction is done
    def close_application(self):
        """Close the application"""
        excel_process() # Update the Excel file clendar
        self.log.emit("👋 Goodbye!")
        QApplication.quit()
        self.finished.emit()
            
    # This method runs the check for the updated file in a loop until it finds the file        
    def run_download_updated_file(self, context):
        '''This method runs the check for the updated 
        file in a loop until it finds the file.'''
        firstCheck = True
        while True:
            self.log.emit(f"🔁 Running check at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            try:
                downloaded_path = run_check(context)
                if downloaded_path:
                    print("✅ File download complete. Exiting loop.")
                    return downloaded_path
            except Exception as e:
                print(f"❗ Error during check: {e}")
            # Wait before the next check
            if firstCheck:
                self.log.emit(f"⏳ Waiting for {FIRST_INTERVAL_SECONDS/3600:.1f} hr before next check...")
                time.sleep(FIRST_INTERVAL_SECONDS)
                firstCheck = False
            else:    
                self.log.emit(f"⏳ Waiting for {SECOND_INTERVAL_SECONDS/3600:.1f} hr before next check...")
                time.sleep(SECOND_INTERVAL_SECONDS)
    

""" Custom Error Overlay Widget"""
class ErrorOverlay(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("error_overlay")
        self.init_ui()
        self.hide()  # Initially hidden
        
    def init_ui(self):
        # Cover half the window
        self.setFixedSize(400, 125)
        
        # Layout
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 15, 20, 15)
        
        # Error icon and title
        title_layout = QHBoxLayout()
        error_icon = QLabel("⚠️")
        error_icon.setFont(QFont("Arial", 20))
        
        title_label = QLabel("IMPORT ERROR")
        title_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        
        title_layout.addWidget(error_icon)
        title_layout.addWidget(title_label)
        title_layout.addStretch()
        
        # Error message
        self.error_message = QLabel("An error occurred during import")
        self.error_message.setObjectName("error_message")
        self.error_message.setWordWrap(True)
        self.error_message.setFont(QFont("Arial", 10))
        
        # Buttons
        button_layout = QHBoxLayout()
        self.close_btn = QPushButton("Close")
        self.close_btn.setObjectName("close_btn")
        self.close_btn.clicked.connect(self.hide_overlay)
        button_layout.addWidget(self.close_btn)
        
        # Add to main layout
        layout.addLayout(title_layout)
        layout.addWidget(self.error_message)
        layout.addStretch()
        layout.addLayout(button_layout)
        
        
    def show_error(self, error_message, error_type="Unknown"):
        """Show the error overlay with the specified message"""
        detailed_message = f"{error_message}\n\nError Type: {error_type}"
        self.error_message.setText(detailed_message)
        
        # Position overlay in center of parent
        if self.parent():
            parent_rect = self.parent().rect()
            x = (parent_rect.width() - self.width()) // 2
            y = (parent_rect.height() - self.height()) // 2
            self.move(x, y)
        
        # Show with animation
        self.show()
        self.raise_()  # Bring to front
        
        # Optional: Auto-hide after 10 seconds
        QTimer.singleShot(10000, self.hide_overlay)
        
    def hide_overlay(self):
        """Hide the error overlay"""
        self.hide()

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Dumhumby Download Export/Import files")
        self.setWindowIcon(QIcon("Dumhumby/logo.png"))
        self.setFixedSize(400, 250)
        self.init_ui()
        self.load_stylesheet()
        
        # Create error overlay
        self.error_overlay = ErrorOverlay(self)

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(20)
        
        # container for the main layout
        container = QWidget()
        container.setObjectName("container_top")
        container_layout = QVBoxLayout(container)
        container_layout.setSpacing(10)
        # Row 1: Download
        download_layout = QHBoxLayout(container)
        download_label = QLabel("Download export files:")
        download_label.setObjectName("title")
        download_button = QPushButton("Download")
        download_button.setObjectName("step_btn")
        download_button.clicked.connect(self.handle_download)

        download_layout.addWidget(download_label)
        download_layout.addStretch()
        download_layout.addWidget(download_button)
        

        # Row 2: Upload/Import
        file_row = QHBoxLayout(container)
        self.file_path_label = QLabel("Upload file: No file selected")
        self.file_path_label.setObjectName("title")
        self.upload_button = QPushButton("Browse")
        self.upload_button.setObjectName("step_btn")
        self.upload_button.clicked.connect(self.select_file)

        file_row.addWidget(self.file_path_label)
        file_row.addStretch()
        file_row.addWidget(self.upload_button)
        
        container_layout.addLayout(download_layout)
        container_layout.addLayout(file_row)
        layout.addWidget(container)
        
        # Status label
        self.status_label = QLabel("Status: Idle")
        self.status_label.setObjectName("status_label")
        self.status_label.setFixedHeight(30)
        layout.addWidget(self.status_label)
        self.setLayout(layout)
        
        # Row 3: Download sales file
        self.sale_style = QWidget()
        self.sale_style.setObjectName("container")
        self.sale_style.setMaximumHeight(60)
        sale_row = QHBoxLayout(self.sale_style)
        sale_row.setAlignment(Qt.AlignmentFlag.AlignVCenter)
        self.sale_row_label = QLabel("To ONLY download sales file:")
        self.sale_row_label.setObjectName("title")
        self.sale_row_button = QPushButton("Download Sales File")
        self.sale_row_button.setToolTip("If close the program and final download didn't complete,\n keep downloading the sale file though here.")
        self.sale_row_button.setObjectName("step3_btn")
        self.sale_row_button.clicked.connect(self.start_download_thread)

        sale_row.addWidget(self.sale_row_label)
        sale_row.addStretch()
        sale_row.addWidget(self.sale_row_button)
        layout.addWidget(self.sale_style)#, alignment=Qt.AlignmentFlag.AlignCenter)


        
    '''===== Funcions to handle button clicks and update status ====='''
    def handle_download(self):
        self.status_label.setText("Status: Starting download...")
        self.download_worker = DownloadWorker()
        self.download_worker.log.connect(self.update_status)
        self.download_worker.finished.connect(self.download_finished)
        self.download_worker.start()
        self.setDisabled(True)
        
    def handle_import(self):
        if not hasattr(self, "selected_file") or not self.selected_file:
            self.update_status("Status: Please select a file to import.")
            return

        self.update_status("Status: Starting import...")
        self.setDisabled(True)
        self.import_worker = ImportAndDownloadFile(self.selected_file)
        self.import_worker.log.connect(self.update_status)
        self.import_worker.import_failed.connect(self.show_import_error)
        self.import_worker.finished.connect(self.start_download_thread)
        self.import_worker.start()

        # Reset to original state
        self.upload_button.setText("Browse")
        self.upload_button.clicked.disconnect()
        self.upload_button.clicked.connect(self.select_file)
        self.file_path_label.setText("Upload file: No file selected")
        self.selected_file = None
    
    def show_import_error(self, error_message, error_type):
        """Show the prominent error overlay for import failures"""
        self.error_overlay.show_error(error_message, error_type)       
        # Re-enable the main window
        self.setDisabled(False)  
    
    def start_download_thread(self):
        self.update_status(f"🔁 Waiting for download to complete...{FIRST_INTERVAL_SECONDS/3600}hr")
        self.downloader = DownloadProductFile()
        self.downloader.log.connect(self.update_status)
        self.downloader.finished.connect(self.download_finished)
        self.downloader.start()  
        
    def download_finished(self):
        self.status_label.setText("Status: Import and download completed.")
        self.setDisabled(False)
        self.file_path_label.setText("No file selected")
        
    def update_status(self, msg):
        self.status_label.setText(f"Status: {msg}")
        
    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select File", "", "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        if file_path:
            self.file_path_label.setText(file_path)
            self.upload_button.setText("Import")
            self.upload_button.clicked.disconnect()
            self.upload_button.clicked.connect(self.handle_import)
            self.selected_file = file_path
        else:
            self.file_path_label.setText("Upload file: No file selected")
            self.upload_button.setText("Browse")
            self.upload_button.clicked.disconnect()
            self.upload_button.clicked.connect(self.select_file)
            self.selected_file = None
    # Load and apply the QSS file
    def load_stylesheet(self):
        # This works whether you're running from source or as a compiled executable
        if hasattr(sys, '_MEIPASS'):
            # Running as compiled executable
            base_path = sys._MEIPASS
        else:
            # Running from source
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        qss_file = os.path.join(base_path, 'styles.qss')
        
        try:
            with open(qss_file, 'r') as file:
                stylesheet = file.read()
                self.setStyleSheet(stylesheet)
                print("Stylesheet loaded successfully")
        except FileNotFoundError:
            print(f"QSS file not found at: {qss_file}")
        except Exception as e:
            print(f"Error loading stylesheet: {e}")


    
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

