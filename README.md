# Excel Monitor and API Sender

A Python application that monitors Excel files for new rows and sends them to a specified API endpoint.

## Features

- Monitor Excel files for new rows
- Send new rows to a configurable API endpoint
- Real-time logging in the UI
- Configurable API URL and API key
- User-friendly interface

## Installation

1. Clone this repository
2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```
   python excel_monitor.py
   ```
2. Configure the API URL and API key in the settings section
3. Select an Excel file to monitor
4. The application will automatically detect and send new rows to the configured API

## Building Executable

To create an executable file:
```
pyinstaller --onefile --windowed excel_monitor.py
```

The executable will be created in the `dist` directory.

## Requirements

- Python 3.8+
- PyQt6
- pandas
- openpyxl
- requests 