# Gestionnaire de Carnet d'Adresses pour RustDesk

Ce projet est un gestionnaire de connexions pour RustDesk, permettant de gérer facilement les connexions à différents appareils à partir d'un carnet d'adresses stocké dans un fichier Excel en local.

## Fonctionnalités

- Chargement des clients et de leurs appareils à partir d'un fichier Excel
- Interface graphique moderne et intuitive
- Sélection facile des clients et des appareils
- Saisie sécurisée des mots de passe
- Connexion rapide aux appareils distants

### Version Interface Graphique (Recommandée)
Téléchargez et exécutez la version .exe disponible dans "Interface graphique"

### Version Console
1) Si vous préférez utiliser la version console, téléchargez et exécutez la version le "rustdesk.py" disponible dans "CLI" :

2) 
```bash
python rustdesk.py
```

## Configuration

Les 2 versions recherchent l'exécutable RustDesk dans les emplacements suivants :
- `C:\Program Files\RustDesk\rustdesk.exe`
- `C:\Program Files (x86)\RustDesk\rustdesk.exe`

Assurez-vous que RustDesk est installé dans l'un de ces emplacements.

Les 2 versions utilisent le carnet d'adresses suivant :
- `C:\Windows\address_book.xlsx`

Assurez-vous que le carnet d'adresses est disponible dans cet emplacement sous le bon format de colonnes : 
- `Client`
- `Nom du PC`
- `Identifiant`

## Contribuer

Les contributions sont les bienvenues ! N'hésitez pas à soumettre des demandes de tirage (pull requests) pour des améliorations ou des corrections.