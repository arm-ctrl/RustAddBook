import os
import pandas as pd
from colorama import init, Fore, Style
from getpass import getpass

# Initialiser colorama
init(autoreset=True)

# Fonction pour afficher un en-tête
def print_header(title):
    print(Fore.WHITE + title)
    print(Fore.YELLOW + "-" * len(title))

# Affichage d'un en-tête au début du script
print_header("Bienvenue dans le gestionnaire de connexions RustDesk")
print(Fore.LIGHTWHITE_EX + "Chargement du carnet d'adresses...\n")

# Fonction pour lire le carnet d'adresses
def load_address_book(file_path):
    try:
        # Lire le fichier Excel
        df = pd.read_excel(file_path)
        # Vérifier que toutes les colonnes nécessaires sont présentes
        required_columns = ['Client', 'Nom du PC', 'Identifiant']
        if not all(col in df.columns for col in required_columns):
            print("Format du fichier Excel incorrect. Les colonnes requises sont:", required_columns)
            return None
        return df
    except FileNotFoundError:
        print(f"Erreur : Le fichier '{file_path}' n'a pas été trouvé. Veuillez vérifier que le fichier existe à cet emplacement.")
        return None
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier Excel : {str(e)}")
        return None

# Fonction pour connecter à un appareil
def connect_to_device(rustdesk_path, device_id, password):
    command = f'"{rustdesk_path}" --connect {device_id} --password {password}'
    os.system(command)

# Fonction pour afficher les appareils par groupe
def display_grouped_devices(address_book):
    # Grouper les appareils par client
    grouped = address_book.groupby('Client')
    
    # Dictionnaire pour stocker les appareils par index
    devices_dict = {}
    current_index = 1
    
    print("\n" + Fore.GREEN + "Liste des appareils par client:")
    print(Fore.YELLOW + "-" * 50)
    
    for client_name, group in grouped:
        print("\n" + Fore.LIGHTYELLOW_EX + f"Client: {client_name}")
        print(Fore.YELLOW + "-" * 20)
        
        for _, row in group.iterrows():
            print(f"{current_index}. {row['Nom du PC']} (ID: {row['Identifiant']})")
            devices_dict[current_index] = row
            current_index += 1
        
    return devices_dict

# Chemins possibles pour l'exécutable RustDesk
RUSTDESK_PATH_1 = r"C:\Program Files\RustDesk\rustdesk.exe"
RUSTDESK_PATH_2 = r"C:\Program Files (x86)\RustDesk\rustdesk.exe"

# Vérifier quel chemin existe
if os.path.exists(RUSTDESK_PATH_1):
    RUSTDESK_PATH = RUSTDESK_PATH_1
elif os.path.exists(RUSTDESK_PATH_2):
    RUSTDESK_PATH = RUSTDESK_PATH_2
else:
    print("Aucun chemin valide trouvé pour RustDesk.")
    RUSTDESK_PATH = None

# Chemin du fichier Excel où il y a le carnet d'adresses
ADDRESS_BOOK_FILE = r"C:\Windows\address_book.xlsx"

# Chargement du carnet d'adresses
address_book = load_address_book(ADDRESS_BOOK_FILE)

if address_book is not None:
    print(Fore.GREEN + "Carnet d'adresses chargé avec succès !")
    
    # Boucle principale
    while True:
        # Afficher les appareils groupés et obtenir le dictionnaire des appareils
        devices_dict = display_grouped_devices(address_book)
        
        # Demander à l'utilisateur de choisir un appareil
        choice = input("\n" + Fore.CYAN + "Entrez le numéro de l'appareil à connecter (ou 'q' pour quitter) : ")
        
        if choice.lower() == 'q':
            break
        
        try:
            index = int(choice)
            if index in devices_dict:
                device = devices_dict[index]
                if RUSTDESK_PATH is not None:
                    # Demander le mot de passe à l'utilisateur avec getpass
                    print(Fore.LIGHTMAGENTA_EX + "Veuillez entrer le mot de passe pour l'appareil sélectionné : ", end='')
                    password = getpass('')
                    connect_to_device(RUSTDESK_PATH, device['Identifiant'], password)
                else:
                    print("Impossible de se connecter, aucun chemin valide trouvé pour RustDesk.")
            else:
                print(Fore.RED + "Numéro d'appareil invalide. Veuillez réessayer.")
        except ValueError:
            print(Fore.RED + "Entrée invalide. Veuillez entrer un numéro ou 'q' pour quitter.")