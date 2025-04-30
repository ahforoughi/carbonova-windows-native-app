import os
import time
import pandas as pd
import requests
from datetime import datetime
from PyQt5.QtCore import QThread, pyqtSignal

class FileMonitor(QThread):
    log_signal = pyqtSignal(str)
    
    def __init__(self, file_to_monitor, api_url, api_key):
        super().__init__()
        self.file_to_monitor = file_to_monitor
        self.api_url = api_url
        self.api_key = api_key
        self.running = True
        self.file_state = {'row_count': 0, 'mod_time': 0}
        self.check_interval = 2  # seconds between file checks
        
    def run(self):
        self.log_signal.emit(f"Started monitoring file: {os.path.basename(self.file_to_monitor)}")
        
        # Initial scan of file
        self.process_file(initial_scan=True)
        
        try:
            while self.running:
                if os.path.exists(self.file_to_monitor):
                    self.process_file()
                else:
                    self.log_signal.emit(f"Warning: File not found: {self.file_to_monitor}")
                
                time.sleep(self.check_interval)
                
        except Exception as e:
            self.log_signal.emit(f"Error in monitoring thread: {str(e)}")
            
    def process_file(self, initial_scan=False):
        try:
            # Check if file exists
            if not os.path.exists(self.file_to_monitor):
                self.log_signal.emit(f"Warning: File not found: {self.file_to_monitor}")
                return
            
            # Check file modification time
            mod_time = os.path.getmtime(self.file_to_monitor)
            
            # Skip if we've already processed this file with this mod time
            if mod_time <= self.file_state.get('mod_time', 0) and not initial_scan:
                return  # File hasn't changed
            
            # Initialize DataFrame variable
            df = None
            
            # First try to read the file based on its extension
            try:
                if self.file_to_monitor.endswith('.csv'):
                    df = pd.read_csv(self.file_to_monitor)
                elif self.file_to_monitor.endswith('.xls'):
                    df = pd.read_excel(self.file_to_monitor, engine='xlrd')
                elif self.file_to_monitor.endswith('.xlsx'):
                    df = pd.read_excel(self.file_to_monitor, engine='openpyxl')
                else:
                    df = pd.read_excel(self.file_to_monitor)
            except Exception as excel_error:
                # If standard approach fails, try alternative methods
                self.log_signal.emit(f"Standard reading failed, trying alternative methods: {str(excel_error)}")
                
                # Check if it's actually a text file with wrong extension
                try:
                    # Try reading as tab-delimited file
                    df = pd.read_csv(self.file_to_monitor, sep='\t')
                    self.log_signal.emit(f"Successfully read file as tab-delimited file")
                except Exception:
                    try:
                        # Try reading as CSV
                        df = pd.read_csv(self.file_to_monitor)
                        self.log_signal.emit(f"Successfully read file as CSV")
                    except Exception as final_error:
                        # All attempts failed
                        raise Exception(f"Failed to read file with any method: {str(final_error)}")
            
            if df is None:
                raise Exception("Failed to load the file with any available method")
            
            # Get the current row count
            row_count = len(df)
            
            # If we've seen this file before
            previous_count = self.file_state.get('row_count', 0)
            
            # If new rows were added
            if row_count > previous_count:
                new_rows = df.iloc[previous_count:row_count]
                self.log_signal.emit(f"Found {len(new_rows)} new rows")
                
                # Send new rows to API
                if not initial_scan:
                    self.send_to_api(new_rows)
            
            # Update the file state
            self.file_state['row_count'] = row_count
            self.file_state['mod_time'] = mod_time
            
            if initial_scan:
                self.log_signal.emit(f"Initially indexed file with {row_count} rows")
                
        except Exception as e:
            self.log_signal.emit(f"Error processing file: {str(e)}")
            
            # Provide guidance based on the error
            if "Expected BOF record" in str(e):
                self.log_signal.emit(f"The file appears to have an Excel extension but is not a valid Excel file.")
                self.log_signal.emit(f"It might be a text or CSV file with an incorrect extension.")
    
    def send_to_api(self, data):
        try:
            # Convert DataFrame to dict for JSON serialization
            records = data.to_dict(orient='records')
            
            # Prepare the payload
            payload = {
                'file_name': os.path.basename(self.file_to_monitor),
                'timestamp': datetime.now().isoformat(),
                'data': records
            }
            
            # Prepare headers with API key
            headers = {
                'Content-Type': 'application/json',
                'Authorization': f'Bearer {self.api_key}'
            }
            
            # Send the request
            response = requests.post(self.api_url, json=payload, headers=headers)
            
            # Log the result
            if response.status_code == 200:
                self.log_signal.emit(f"Successfully sent {len(records)} rows to API")
            else:
                self.log_signal.emit(f"API error: {response.status_code} - {response.text}")
                
        except Exception as e:
            self.log_signal.emit(f"Error sending data to API: {str(e)}")
    
    def stop(self):
        self.running = False