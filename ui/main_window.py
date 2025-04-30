import os
from datetime import datetime
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLabel, QLineEdit, QTextEdit, 
                             QFileDialog, QTabWidget, QGroupBox, QFormLayout, 
                             QMessageBox, QTableWidget, QTableWidgetItem, QSplitter,
                             QHeaderView)
from PyQt5.QtCore import Qt
from settings.settings_manager import SettingsManager
from monitoring.file_monitor import FileMonitor

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel File Monitor")
        
        # Set a larger initial window size
        self.setMinimumSize(1000, 700)
        self.resize(1200, 800)
        
        self.settings_manager = SettingsManager()
        self.monitor_thread = None
        self.file_to_monitor = ""
        
        self.init_ui()
        self.load_settings()
        
    def init_ui(self):
        # Main widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Create tabs
        tabs = QTabWidget()
        main_layout.addWidget(tabs)
        
        # Monitor tab
        monitor_tab = QWidget()
        monitor_layout = QVBoxLayout(monitor_tab)
        
        # Status group
        status_group = QGroupBox("Status")
        status_layout = QVBoxLayout(status_group)
        
        self.status_label = QLabel("Monitoring: Inactive")
        status_layout.addWidget(self.status_label)
        
        # Currently monitoring label
        self.current_file_label = QLabel("File: None")
        status_layout.addWidget(self.current_file_label)
        
        # Control buttons
        control_layout = QHBoxLayout()
        self.start_button = QPushButton("Start Monitoring")
        self.start_button.clicked.connect(self.start_monitoring)
        self.stop_button = QPushButton("Stop Monitoring")
        self.stop_button.clicked.connect(self.stop_monitoring)
        self.stop_button.setEnabled(False)
        
        control_layout.addWidget(self.start_button)
        control_layout.addWidget(self.stop_button)
        status_layout.addLayout(control_layout)
        
        monitor_layout.addWidget(status_group)
        
        # Create a splitter for log and data view
        splitter = QSplitter(Qt.Vertical)
        
        # Log area
        log_group = QGroupBox("Activity Log")
        log_layout = QVBoxLayout(log_group)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        splitter.addWidget(log_group)
        
        # Data view area - create as a class member
        self.data_group = QGroupBox("New Rows Added (0)")
        data_layout = QVBoxLayout(self.data_group)
        
        self.data_table = QTableWidget()
        self.data_table.setAlternatingRowColors(True)
        data_layout.addWidget(self.data_table)
        
        # Initialize with empty table
        self.data_table.setRowCount(0)
        self.data_table.setColumnCount(0)
        
        splitter.addWidget(self.data_group)
        
        # Set initial sizes for the splitter
        splitter.setSizes([200, 400])
        
        monitor_layout.addWidget(splitter)
        
        # Settings tab
        settings_tab = QWidget()
        settings_layout = QVBoxLayout(settings_tab)
        
        form_layout = QFormLayout()
        
        # File selection
        file_layout = QHBoxLayout()
        self.file_path_input = QLineEdit()
        self.file_path_input.setReadOnly(True)
        file_browse_button = QPushButton("Browse...")
        file_browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(self.file_path_input)
        file_layout.addWidget(file_browse_button)
        form_layout.addRow("File to Monitor:", file_layout)
        
        # API URL
        self.api_url_input = QLineEdit()
        form_layout.addRow("API URL:", self.api_url_input)
        
        # API Key
        self.api_key_input = QLineEdit()
        form_layout.addRow("API Key:", self.api_key_input)
        
        settings_layout.addLayout(form_layout)
        
        # Save button
        save_layout = QHBoxLayout()
        save_button = QPushButton("Save Settings")
        save_button.clicked.connect(self.save_settings)
        save_layout.addStretch()
        save_layout.addWidget(save_button)
        settings_layout.addLayout(save_layout)
        
        settings_layout.addStretch()
        
        # Add tabs
        tabs.addTab(monitor_tab, "Monitor")
        tabs.addTab(settings_tab, "Settings")
        
    def load_settings(self):
        settings = self.settings_manager.settings
        self.file_to_monitor = settings.get('file_to_monitor', '')
        self.file_path_input.setText(self.file_to_monitor)
        self.api_url_input.setText(settings.get('api_url', ''))
        self.api_key_input.setText(settings.get('api_key', ''))
        
        # Update current file label
        if self.file_to_monitor:
            self.current_file_label.setText(f"File: {os.path.basename(self.file_to_monitor)}")
        else:
            self.current_file_label.setText("File: None")
        
    def save_settings(self):
        settings = {
            'file_to_monitor': self.file_to_monitor,
            'api_url': self.api_url_input.text().strip(),
            'api_key': self.api_key_input.text().strip()
        }
        
        if not self.file_to_monitor:
            QMessageBox.warning(self, "Invalid Settings", "Please select a file to monitor.")
            return
            
        if not settings['api_url']:
            QMessageBox.warning(self, "Invalid Settings", "Please enter a valid API URL.")
            return
            
        if self.settings_manager.save_settings(settings):
            QMessageBox.information(self, "Settings Saved", "Your settings have been saved successfully.")
        else:
            QMessageBox.warning(self, "Error", "Failed to save settings.")
            
    def browse_file(self):
        file, _ = QFileDialog.getOpenFileName(
            self, 
            "Select File to Monitor", 
            "", 
            "Excel Files (*.xlsx *.xls *.csv)"
        )
        
        if file:
            self.file_to_monitor = file
            self.file_path_input.setText(file)
            self.current_file_label.setText(f"File: {os.path.basename(file)}")
            
    def log_message(self, message):
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.log_text.append(f"[{timestamp}] {message}")
        
    def display_new_data(self, dataframe):
        if dataframe is None or len(dataframe) == 0:
            self.log_message("No new data to display")
            return
        
        # Debug message to confirm data is being received
        self.log_message(f"Displaying {len(dataframe)} new rows in table")
        
        # Clear the table
        self.data_table.clear()
        
        # Set the number of rows and columns
        row_count = len(dataframe)
        col_count = len(dataframe.columns)
        
        self.data_table.setRowCount(row_count)
        self.data_table.setColumnCount(col_count)
        
        # Set the headers
        self.data_table.setHorizontalHeaderLabels(dataframe.columns)
        
        # Populate the table with data
        for row in range(row_count):
            for col in range(col_count):
                value = str(dataframe.iloc[row, col])
                item = QTableWidgetItem(value)
                self.data_table.setItem(row, col, item)
        
        # Resize columns to content
        self.data_table.resizeColumnsToContents()
        
        # Update the group box title to show it's only new rows
        self.data_group.setTitle(f"New Rows Added ({row_count})")
        
    def start_monitoring(self):
        # Check if already running
        if self.monitor_thread and self.monitor_thread.isRunning():
            return
            
        # Get settings
        if not self.file_to_monitor:
            QMessageBox.warning(self, "Invalid Settings", "Please select a file to monitor.")
            return
            
        api_url = self.api_url_input.text().strip()
        api_key = self.api_key_input.text().strip()
        
        # Validate settings
        if not api_url:
            QMessageBox.warning(self, "Invalid Settings", "Please enter a valid API URL.")
            return
            
        # Reset the data view
        self.data_table.setRowCount(0)
        self.data_table.setColumnCount(0)
        self.data_group.setTitle("New Rows Added (0)")
            
        # Start monitoring thread
        self.monitor_thread = FileMonitor(self.file_to_monitor, api_url, api_key)
        self.monitor_thread.log_signal.connect(self.log_message)
        self.monitor_thread.new_data_signal.connect(self.display_new_data)
        self.monitor_thread.start()
        
        # Update UI
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.status_label.setText("Monitoring: Active")
        
        self.log_message(f"Monitoring started for {os.path.basename(self.file_to_monitor)}")
        
    def stop_monitoring(self):
        if self.monitor_thread and self.monitor_thread.isRunning():
            self.monitor_thread.stop()
            self.monitor_thread.wait()
            
            # Update UI
            self.start_button.setEnabled(True)
            self.stop_button.setEnabled(False)
            self.status_label.setText("Monitoring: Inactive")
            
            self.log_message("Monitoring stopped")
            
    def closeEvent(self, event):
        # Stop the monitoring thread when the application closes
        if self.monitor_thread and self.monitor_thread.isRunning():
            self.monitor_thread.stop()
            self.monitor_thread.wait()
        event.accept()