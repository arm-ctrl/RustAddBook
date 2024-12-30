# Interface RustDesk

Une interface graphique moderne pour RustDesk permettant de gérer les connexions à distance et de suivre l'historique des connexions.

## Fonctionnalités

- Interface graphique moderne avec thème sombre
- Gestion des connexions à distance via RustDesk
- Carnet d'adresses des clients et appareils
- Historique des connexions avec filtrage par date et heure
- Statistiques sur les connexions (nombre total, clients uniques, appareils)
- Export des données vers Excel
- Gestion automatique des logs de connexion

## Prérequis

- Python 3
- RustDesk installé sur le système et configuré avec le Server ID et le Key.
- Fichier `address_book.xlsx` dans `C:\Windows\` (format requis : colonnes 'Client', 'Nom du PC', 'Identifiant')

## Installation des dépendances

```bash
pip install customtkinter pandas openpyxl numpy
```

## Utilisation

### Depuis le code source
```bash
python rustdesk_modern.py
```

Si vous effectuez une mise à jour dans le code, vous pouvez utiliser la commande suivante pour regénerer l'exécutable :
```bash
pyinstaller --onefile --windowed --clean rustdesk_modern.py
```

### Version exécutable
Téléchargez la dernière version de l'exécutable et lancez-le directement.

## Structure des fichiers

- `rustdesk_modern.py` : Application principale
- `RustDesk Connector.exe` : Application avec interface graphique

## Fonctionnement

1. L'application vérifie la présence du carnet d'adresses et de RustDesk
2. Les clients sont affichés dans la colonne de gauche
3. Sélectionnez un client pour voir ses appareils
4. Cliquez sur un appareil, puis saisissez le mot de passe pour initier une connexion
5. L'historique des connexions est automatiquement enregistré

## Stockage des données

- Carnet d'adresses : `C:\Windows\address_book.xlsx`
- Historique des connexions : 
  - Windows : `%APPDATA%\RustDeskInterface\logs\connection_history.json`
  - macOS/Linux : `~/.rustdeskinterface/logs/connection_history.json`

## Notes

- L'interface utilise customtkinter pour un design moderne
- Les logs sont automatiquement gérés et peuvent être exportés
- L'historique peut être filtré par date et heure
- Les statistiques sont mises à jour en temps réel