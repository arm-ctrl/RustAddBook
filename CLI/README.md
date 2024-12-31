# RustDesk CLI

A command-line interface for RustDesk, allowing users to manage remote connections easily.

## Features

- Simple and efficient command-line interface
- Remote connection management via RustDesk
- Reading of Excel address book
- Colorful display for better readability
- Password protection for certain features

## Prerequisites

- Python 3
- RustDesk installed on the system
- `address_book.xlsx` file in `C:\Windows\` (required format: columns 'Client', 'Hostname', 'Rustdesk_ID')

## Installing Dependencies

```bash
pip install pandas colorama openpyxl
```

## Usage

```bash
python rustdesk.py
```

## File Structure

- `rustdesk.py` : Main command-line script

## Functioning

1. The script checks for the presence of the address book
2. List of clients displayed with color code
3. Client selection by number
4. Choice of device to connect to
5. Automatic launch of RustDesk connection

## Notes

- Simple usage via command-line
- Colorful display with colorama for better readability
- Password protection of sensitive functions
- Command-line version, complement to the modern graphical version