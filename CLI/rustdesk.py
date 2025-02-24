import os
import pandas as pd
from colorama import init, Fore, Style
from getpass import getpass
import sys

# Initialize colorama
init(autoreset=True)

# Function to display a header
def print_header(title):
    print(Fore.WHITE + title)
    print(Fore.YELLOW + "-" * len(title))

# Display a header at the beginning of the script
print_header("Welcome to the RustDesk Connection Manager")
print(Fore.LIGHTWHITE_EX + "Loading address book...\n")

# Constants
ADDRESS_BOOK_FILE = r"C:\Windows\address_book.csv"

# Function to read the address book
def read_address_book(file_path):
    """Read the address book from CSV file."""
    try:
        df = pd.read_csv(file_path)
        # Check that all required columns are present
        required_columns = ['Client', 'Hostname', 'Rustdesk_ID']
        if not all(col in df.columns for col in required_columns):
            print("Incorrect CSV file format. Required columns are:", required_columns)
            return None
        return df
    except FileNotFoundError:
        print(f"Error: Address book not found at {file_path}")
        sys.exit(1)
    except Exception as e:
        print(f"Error reading address book: {str(e)}")
        sys.exit(1)

# Function to connect to a device
def connect_to_device(rustdesk_path, device_id, password):
    # Use 'start' for Windows to run RustDesk in the background
    command = f'start "" "{rustdesk_path}" --connect {device_id} --password {password}'
    os.system(command)
    print(Fore.GREEN + "\nConnection initiated successfully!")

# Function to display devices by group
def display_grouped_devices(address_book):
    # Group devices by client
    grouped = address_book.groupby('Client')
    
    # Dictionary to store devices by index
    devices_dict = {}
    current_index = 1
    
    print("\n" + Fore.GREEN + "List of devices by client:")
    print(Fore.YELLOW + "-" * 50)
    
    for client_name, group in grouped:
        print("\n" + Fore.LIGHTYELLOW_EX + f"Client: {client_name}")
        print(Fore.YELLOW + "-" * 20)
        
        for _, row in group.iterrows():
            print(f"{current_index}. {row['Hostname']} (ID: {row['Rustdesk_ID']})")
            devices_dict[current_index] = row
            current_index += 1
        
    return devices_dict

# Possible paths for the RustDesk executable
RUSTDESK_PATH_1 = r"C:\Program Files\RustDesk\rustdesk.exe"
RUSTDESK_PATH_2 = r"C:\Program Files (x86)\RustDesk\rustdesk.exe"

# Check which path exists
if os.path.exists(RUSTDESK_PATH_1):
    RUSTDESK_PATH = RUSTDESK_PATH_1
elif os.path.exists(RUSTDESK_PATH_2):
    RUSTDESK_PATH = RUSTDESK_PATH_2
else:
    print("No valid path found for RustDesk.")
    RUSTDESK_PATH = None

# Load the address book
address_book = read_address_book(ADDRESS_BOOK_FILE)

if address_book is not None:
    print(Fore.GREEN + "Address book loaded successfully!")
    
    # Main loop
    while True:
        # Display grouped devices and get the device dictionary
        devices_dict = display_grouped_devices(address_book)
        
        # Ask the user to choose a device
        choice = input("\n" + Fore.CYAN + "Enter the number of the device to connect to (or 'q' to quit): ")
        
        if choice.lower() == 'q':
            break
        
        try:
            index = int(choice)
            if index in devices_dict:
                device = devices_dict[index]
                # Loop to ask for the password until it is valid
                while True:
                    password = getpass(Fore.LIGHTMAGENTA_EX + "Please enter the password for the selected device: ")
                    if not password.strip():
                        print(Fore.RED + "The password cannot be empty!")
                        continue
                    break
                
                connect_to_device(RUSTDESK_PATH, device['Rustdesk_ID'], password)  # Connect to the device
                
                # Ask if the user wants to reconnect
                reconnect = input("Do you want to reconnect to another device? (Y/n): ") or 'y'
                if reconnect.lower() != 'y':
                    break
            else:
                print(Fore.RED + "Invalid device number. Please try again.")
        except ValueError:
            print(Fore.RED + "Invalid input. Please enter a number or 'q' to quit.")