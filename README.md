# RustAddBook

A modern graphical interface for RustDesk to manage remote connections and track connection history.

## Features

- Modern graphical interface with dark theme
- Remote connection management via RustDesk
- Client and device address book
- Connection history with date and time filtering
- Connection statistics (total count, unique clients, devices)
- Data export to Excel
- Automatic connection logging
- Multi-language support (English, French, Spanish)

## Prerequisites

- Python 3
- RustDesk installed and configured with Server ID and Key
- `address_book.xlsx` file in `C:\Windows\` (required format: columns 'Client', 'Hostname', 'Rustdesk_ID')

## Installing Dependencies

```bash
pip install customtkinter pandas openpyxl numpy
```

## Usage

### From Source Code
```bash
python rustdesk_modern.py
```

If you make changes to the code, you can regenerate the executable using:
```bash
pyinstaller --onefile --windowed --clean rustdesk_modern.py
```

### Executable Version
Download the latest version of the executable and run it directly.

## File Structure

- `rustdesk_modern.py`: Main application
- `RustDesk Connector.exe`: GUI application

## How It Works

1. The application checks for the presence of the address book and RustDesk
2. Clients are displayed in the left column
3. Select a client to view their devices
4. Click on a device, then enter the password to initiate a connection
5. Connection history is automatically logged

## Notes

- The interface uses customtkinter for a modern design
- Logs are automatically managed and can be exported
- History can be filtered by date and time
- Statistics are updated in real-time
- Multi-language support (English, French, Spanish)