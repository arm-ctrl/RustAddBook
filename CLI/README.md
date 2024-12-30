# RustDesk CLI

Une interface en ligne de commande pour RustDesk permettant de gérer facilement les connexions à distance.

## Fonctionnalités

- Interface en ligne de commande simple et efficace
- Gestion des connexions à distance via RustDesk
- Lecture du carnet d'adresses Excel
- Affichage coloré pour une meilleure lisibilité
- Protection par mot de passe pour certaines fonctionnalités

## Prérequis

- Python 3
- RustDesk installé sur le système
- Fichier `address_book.xlsx` dans `C:\Windows\` (format requis : colonnes 'Client', 'Nom du PC', 'Identifiant')

## Installation des dépendances

```bash
pip install pandas colorama openpyxl
```

## Utilisation

```bash
python rustdesk.py
```

## Structure des fichiers

- `rustdesk.py` : Script principal en ligne de commande

## Fonctionnement

1. Le script vérifie la présence du carnet d'adresses
2. Liste des clients affichée avec un code couleur
3. Sélection du client par numéro
4. Choix de l'appareil à connecter
5. Lancement automatique de la connexion RustDesk

## Stockage des données

- Carnet d'adresses : `C:\Windows\address_book.xlsx`
  - Format : Excel avec colonnes 'Client', 'Nom du PC', 'Identifiant'

## Notes

- Utilisation simple via ligne de commande
- Affichage coloré avec colorama pour une meilleure lisibilité
- Protection par mot de passe des fonctions sensibles
- Version en ligne de commande, complément de la version graphique moderne
