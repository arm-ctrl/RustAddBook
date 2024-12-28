# Gestionnaire de Carnet d'Adresses pour RustDesk

Ce projet est un gestionnaire de connexions pour RustDesk, permettant de gérer facilement les connexions à différents appareils à partir d'un carnet d'adresses stocké dans un fichier Excel en local.

## Fonctionnalités

- Chargement des clients et de leurs appareils à partir d'un fichier Excel.
- Interface en ligne de commande pour sélectionner un client et se connecter à un appareil.
- Saisie sécurisée des mots de passe lors de la connexion.

## Prérequis

- Python 3
- Bibliothèques Python nécessaires :
  - pandas
  - colorama
  - getpass (intégré à Python, pas besoin d'installation)

## Installation et Utilisation

1. Télécharger le script python.

2. Ouvrir un CMD, accéder au répertoire où le script est situé :
   ```bash
   cd repo
   ```

3. Installer les dépendances :
   ```bash
   pip install pandas colorama
   ```
   ```bash
   pip install openpyxl
   ```

4. Créer le fichier `address_book.xlsx` dans le répertoire `C:\Windows\` :
   Le format a respecter est le suivant :
   - Colonne 1 : Client
   - Colonne 2 : Nom du PC
   - Colonne 3 : Identifiant

5. Exécuter le script :
   ```bash
   python rustdesk.py
   ```

6. Suivre les instructions affichées dans le terminal.

## Autres

Installer RustDesk en dur sur le PC et ne pas utiliser la version portable.

Le script ira chercher l'executable de RustDesk dans les deux chemins suivants :
- `C:\Program Files\RustDesk\rustdesk.exe`
- `C:\Program Files (x86)\RustDesk\rustdesk.exe`
Il faut donc s'assurer que RustDesk est installé dans l'un de ces 2 emplacements.

## Contribuer

Les contributions sont les bienvenues ! N'hésitez pas à soumettre des demandes de tirage (pull requests) pour des améliorations ou des corrections.