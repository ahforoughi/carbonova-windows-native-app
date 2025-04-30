import sys
import os
import json
import time
import requests
import pandas as pd
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QTextEdit, QLabel, 
                            QLineEdit, QFileDialog, QMessageBox)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

class ExcelFileHandler(FileSystemEventHandler):
    def __init__(self, callback):
        self.callback = callback
        self.last_modified = 0
        self.last_size = 0

    def on_modified(self, event):
        if not event.is_directory and (event.src_path.endswith('.xls') or event.src_path.endswith('.xlsx')):
            current_time = time.time()
            try:
                current_size = os.path.getsize(event.src_path)
                if current_size != self.last_size:
                    self.last_size = current_size
                    if current_time - self.last_modified > 1:  # Debounce for 1 second
                        self.last_modified = current_time
                        self.callback()
            except Exception as e:
                print(f"Error checking file size: {str(e)}")

class MonitorThread(QThread):
    file_changed = pyqtSignal()

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path
        self.observer = None
        self.handler = None

    def run(self):
        directory = os.path.dirname(self.file_path)
        self.handler = ExcelFileHandler(self.on_file_changed)
        self.observer = Observer()
        self.observer.schedule(self.handler, directory, recursive=False)
        self.observer.start()
        try:
            while True:
                time.sleep(1)
        except:
            self.observer.stop()
        self.observer.join()

    def on_file_changed(self):
        self.file_changed.emit()

    def stop(self):
        if self.observer:
            self.observer.stop()
            self.observer.join()

class ExcelMonitorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Monitor and API Sender")
        self.setGeometry(100, 100, 800, 600)
        
        # Initialize variables
        self.excel_file = None
        self.last_row_count = 0
        self.last_columns = None
        self.api_url = ""
        self.api_key = ""
        self.monitoring = False
        self.monitor_thread = None
        self.file_type = None  # 'xls' or 'xlsx'
        
        # Create main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        # Settings section
        settings_group = QWidget()
        settings_layout = QVBoxLayout(settings_group)
        
        # API URL input
        url_layout = QHBoxLayout()
        url_label = QLabel("API URL:")
        self.url_input = QLineEdit()
        url_layout.addWidget(url_label)
        url_layout.addWidget(self.url_input)
        settings_layout.addLayout(url_layout)
        
        # API Key input
        key_layout = QHBoxLayout()
        key_label = QLabel("API Key:")
        self.key_input = QLineEdit()
        self.key_input.setEchoMode(QLineEdit.EchoMode.Password)
        key_layout.addWidget(key_label)
        key_layout.addWidget(self.key_input)
        settings_layout.addLayout(key_layout)
        
        # File selection
        file_layout = QHBoxLayout()
        self.file_label = QLabel("No file selected")
        select_file_btn = QPushButton("Select Excel File")
        select_file_btn.clicked.connect(self.select_file)
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(select_file_btn)
        settings_layout.addLayout(file_layout)
        
        # Control buttons
        control_layout = QHBoxLayout()
        self.start_btn = QPushButton("Start Monitoring")
        self.start_btn.clicked.connect(self.toggle_monitoring)
        self.start_btn.setEnabled(False)
        control_layout.addWidget(self.start_btn)
        settings_layout.addLayout(control_layout)
        
        layout.addWidget(settings_group)
        
        # Log section
        log_label = QLabel("Logs:")
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(log_label)
        layout.addWidget(self.log_text)
        
        # Load saved settings
        self.load_settings()
        
    def log_message(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
        
    def read_excel_file(self, file_path):
        try:
            # First check if the file is actually a text file
            with open(file_path, 'r', encoding='utf-8') as f:
                first_line = f.readline().strip()
                if '\t' in first_line:
                    self.log_message("Detected tab-separated text file")
                    return pd.read_csv(file_path, sep='\t', encoding='utf-8')
            
            # If not a text file, proceed with Excel reading
            if file_path.endswith('.xls'):
                try:
                    return pd.read_excel(file_path, engine='xlrd')
                except Exception as e:
                    self.log_message(f"xlrd failed: {str(e)}")
                    try:
                        return pd.read_excel(file_path, engine='openpyxl')
                    except Exception as e:
                        self.log_message(f"openpyxl failed: {str(e)}")
                        raise ValueError("Could not read the file. Please ensure it is a valid Excel file or tab-separated text file.")
            elif file_path.endswith('.xlsx'):
                return pd.read_excel(file_path, engine='openpyxl')
            else:
                raise ValueError("Unsupported file format")
        except Exception as e:
            self.log_message(f"Error reading file: {str(e)}")
            raise
        
    def select_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xls *.xlsx);;Excel 97-2003 (*.xls);;Excel 2007+ (*.xlsx)"
        )
        if file_name:
            if not (file_name.endswith('.xls') or file_name.endswith('.xlsx')):
                QMessageBox.warning(self, "Error", "Please select an Excel file (.xls or .xlsx)")
                return
                
            self.excel_file = file_name
            self.file_type = 'xls' if file_name.endswith('.xls') else 'xlsx'
            self.file_label.setText(os.path.basename(file_name))
            self.start_btn.setEnabled(True)
            self.log_message(f"Selected file: {file_name}")
            
            try:
                df = self.read_excel_file(self.excel_file)
                self.last_row_count = len(df)
                self.last_columns = df.columns.tolist()
                self.log_message(f"Initial row count: {self.last_row_count}")
                self.log_message(f"Columns: {', '.join(self.last_columns)}")
            except Exception as e:
                self.log_message(f"Error reading initial file: {str(e)}")
                QMessageBox.warning(self, "Error", f"Error reading file: {str(e)}")
                self.start_btn.setEnabled(False)
            
    def toggle_monitoring(self):
        if not self.monitoring:
            if not self.validate_settings():
                return
            self.monitoring = True
            self.start_btn.setText("Stop Monitoring")
            self.start_monitoring()
            self.log_message("Started monitoring")
        else:
            self.monitoring = False
            self.start_btn.setText("Start Monitoring")
            self.stop_monitoring()
            self.log_message("Stopped monitoring")
            
    def start_monitoring(self):
        if self.excel_file and not self.monitor_thread:
            self.monitor_thread = MonitorThread(self.excel_file)
            self.monitor_thread.file_changed.connect(self.check_excel_file)
            self.monitor_thread.start()
            
    def stop_monitoring(self):
        if self.monitor_thread:
            self.monitor_thread.stop()
            self.monitor_thread = None
            
    def validate_settings(self):
        if not self.excel_file:
            QMessageBox.warning(self, "Error", "Please select an Excel file")
            return False
        if not self.url_input.text():
            QMessageBox.warning(self, "Error", "Please enter API URL")
            return False
        if not self.key_input.text():
            QMessageBox.warning(self, "Error", "Please enter API Key")
            return False
        return True
        
    def check_excel_file(self):
        try:
            self.log_message("Checking file for changes...")
            df = self.read_excel_file(self.excel_file)
            current_row_count = len(df)
            current_columns = df.columns.tolist()
            
            self.log_message(f"Current row count: {current_row_count}")
            self.log_message(f"Current columns: {', '.join(current_columns)}")
            
            if self.last_columns != current_columns:
                self.log_message("Warning: Column structure has changed!")
                self.last_columns = current_columns
            
            if current_row_count > self.last_row_count:
                new_rows = df.iloc[self.last_row_count:]
                self.log_message(f"Found {len(new_rows)} new rows")
                self.log_message("New rows data:")
                for idx, row in new_rows.iterrows():
                    self.log_message(f"Row {idx + 1}: {row.to_dict()}")
                
                self.send_to_api(new_rows)
                self.last_row_count = current_row_count
                self.log_message(f"Updated last row count to: {self.last_row_count}")
            else:
                self.log_message("No new rows detected")
                
        except Exception as e:
            self.log_message(f"Error checking file: {str(e)}")
            
    def send_to_api(self, data):
        try:
            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            }
            
            json_data = data.to_json(orient='records')
            
            self.log_message("Sending data to API...")
            response = requests.post(
                self.api_url,
                headers=headers,
                data=json_data
            )
            
            if response.status_code == 200:
                self.log_message("Successfully sent data to API")
            else:
                self.log_message(f"API Error: {response.status_code} - {response.text}")
                
        except Exception as e:
            self.log_message(f"Error sending to API: {str(e)}")
            
    def load_settings(self):
        try:
            if os.path.exists('settings.json'):
                with open('settings.json', 'r') as f:
                    settings = json.load(f)
                    self.url_input.setText(settings.get('api_url', ''))
                    self.key_input.setText(settings.get('api_key', ''))
        except Exception as e:
            self.log_message(f"Error loading settings: {str(e)}")
            
    def save_settings(self):
        try:
            settings = {
                'api_url': self.url_input.text(),
                'api_key': self.key_input.text()
            }
            with open('settings.json', 'w') as f:
                json.dump(settings, f)
        except Exception as e:
            self.log_message(f"Error saving settings: {str(e)}")
            
    def closeEvent(self, event):
        self.stop_monitoring()
        self.save_settings()
        event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelMonitorApp()
    window.show()
    sys.exit(app.exec()) 