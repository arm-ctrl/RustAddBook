import os
import pandas as pd
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from PIL import Image, ImageTk
import sys
from datetime import datetime, timedelta, time
import json
from tkcalendar import DateEntry
import subprocess
import psutil

class ConnectionLogger:
    def __init__(self):
        self.log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
        self.log_file = os.path.join(self.log_dir, 'connection_history.json')
        self.ensure_log_file_exists()
        self.migrate_old_logs()

    def ensure_log_file_exists(self):
        if not os.path.exists(self.log_dir):
            os.makedirs(self.log_dir)
        if not os.path.exists(self.log_file):
            with open(self.log_file, 'w', encoding='utf-8') as f:
                json.dump([], f)

    def migrate_old_logs(self):
        """Convertit les anciennes entrées au nouveau format"""
        try:
            with open(self.log_file, 'r', encoding='utf-8') as f:
                logs = json.load(f)
            
            modified = False
            new_logs = []
            
            for log in logs:
                new_log = {
                    'timestamp': log.get('timestamp', ''),
                    'client': log.get('client', ''),
                    'device_name': log.get('device_name', ''),
                    'device_id': log.get('device_id', ''),
                    'date': log.get('date', ''),
                    'time': log.get('time_connection', log.get('time', ''))  # Prend time_connection ou time s'il existe
                }
                new_logs.append(new_log)
                modified = True
            
            if modified:
                with open(self.log_file, 'w', encoding='utf-8') as f:
                    json.dump(new_logs, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Erreur lors de la migration : {str(e)}")

    def log_connection(self, client, device_name, device_id):
        try:
            with open(self.log_file, 'r', encoding='utf-8') as f:
                logs = json.load(f)
        except:
            logs = []

        connection = {
            'timestamp': datetime.now().isoformat(),
            'client': client,
            'device_name': device_name,
            'device_id': device_id,
            'date': datetime.now().strftime('%d/%m/%Y'),
            'time': datetime.now().strftime('%H:%M')
        }

        logs.append(connection)

        with open(self.log_file, 'w', encoding='utf-8') as f:
            json.dump(logs, f, indent=2, ensure_ascii=False)

    def get_connection_history(self, start_date=None, end_date=None):
        try:
            with open(self.log_file, 'r', encoding='utf-8') as f:
                logs = json.load(f)
            
            if start_date or end_date:
                filtered_logs = []
                for log in logs:
                    try:
                        log_date = datetime.strptime(log['date'], '%d/%m/%Y').date()
                        if start_date and log_date < start_date:
                            continue
                        if end_date and log_date > end_date:
                            continue
                        filtered_logs.append(log)
                    except ValueError:
                        continue
                return filtered_logs
            return logs
        except Exception as e:
            print(f"Erreur lors de la récupération de l'historique : {str(e)}")
            return []

    def delete_old_entries(self, days):
        try:
            with open(self.log_file, 'r', encoding='utf-8') as f:
                logs = json.load(f)
            
            cutoff_date = datetime.now() - timedelta(days=days)
            new_logs = [log for log in logs if datetime.strptime(log['date'], '%d/%m/%Y') > cutoff_date]
            
            with open(self.log_file, 'w', encoding='utf-8') as f:
                json.dump(new_logs, f, indent=2, ensure_ascii=False)
            
            return len(logs) - len(new_logs)  
        except:
            return 0

    def get_statistics(self):
        try:
            with open(self.log_file, 'r', encoding='utf-8') as f:
                logs = json.load(f)
            
            if not logs:
                return {}

            df = pd.DataFrame(logs)
            
            stats = {
                'total_connections': len(logs),
                'unique_clients': len(df['client'].unique()),
                'unique_devices': len(df['device_id'].unique()),
                'most_connected_client': df['client'].mode().iloc[0],
                'most_connected_device': df['device_name'].mode().iloc[0],
                'first_connection': df['timestamp'].min().strftime('%d/%m/%Y %H:%M'),
                'last_connection': df['timestamp'].max().strftime('%d/%m/%Y %H:%M'),
                'connections_by_client': df['client'].value_counts().to_dict(),
                'connections_by_device': df['device_name'].value_counts().to_dict()
            }
            
            return stats
        except:
            return {}

    def get_connection_stats(self):
        try:
            with open(self.log_file, 'r', encoding='utf-8') as f:
                logs = json.load(f)
            
            if not logs:
                return {}

            df = pd.DataFrame(logs)
            
            stats = {
                'total_connections': len(logs),
                'unique_clients': len(df['client'].unique()),
                'unique_devices': len(df['device_id'].unique()),
                'most_connected_client': df['client'].mode().iloc[0],
                'most_connected_device': df['device_name'].mode().iloc[0],
            }
            
            return stats
        except:
            return {}

class HistoryWindow:
    def __init__(self, parent, connection_logger):
        self.window = ctk.CTkToplevel(parent)
        self.window.title("Historique des connexions")
        self.window.geometry("800x600")
        self.window.grab_set()
        
        self.connection_logger = connection_logger
        
        # En-tête avec filtres
        self.create_header()
        
        # Tableau des connexions
        self.create_table()
        
        # Charger les données
        self.load_history()

    def create_header(self):
        header = ctk.CTkFrame(self.window)
        header.pack(fill="x", padx=20, pady=(20, 0))

        # Titre et boutons
        title_frame = ctk.CTkFrame(header, fg_color="transparent")
        title_frame.pack(fill="x", pady=(0, 10))

        # Titre
        ctk.CTkLabel(title_frame, text="Historique des connexions", 
                    font=("Roboto", 20, "bold")).pack(side="left")

        # Boutons d'action
        buttons_frame = ctk.CTkFrame(title_frame, fg_color="transparent")
        buttons_frame.pack(side="right")

        ctk.CTkButton(buttons_frame, text="Exporter Excel", 
                     command=self.export_to_excel,
                     width=100).pack(side="left", padx=5)
                     
        ctk.CTkButton(buttons_frame, text="Rafraîchir", 
                     command=self.load_history,
                     width=100).pack(side="left", padx=5)

        # Statistiques
        stats_frame = ctk.CTkFrame(header, fg_color=("gray90", "gray20"))
        stats_frame.pack(fill="x", pady=(0, 10))

        stats = self.connection_logger.get_connection_stats()
        if stats:
            # Première ligne de stats
            stats_line1 = ctk.CTkFrame(stats_frame, fg_color="transparent")
            stats_line1.pack(fill="x", padx=10, pady=5)
            
            ctk.CTkLabel(stats_line1, 
                        text=f"Total connexions: {stats['total_connections']} • " +
                             f"Clients uniques: {stats['unique_clients']} • " +
                             f"Appareils uniques: {stats['unique_devices']}",
                        font=("Roboto", 12)).pack()

            # Deuxième ligne de stats
            stats_line2 = ctk.CTkFrame(stats_frame, fg_color="transparent")
            stats_line2.pack(fill="x", padx=10, pady=5)
            
            ctk.CTkLabel(stats_line2,
                        text=f"Client le plus connecté: {stats['most_connected_client']} • " +
                             f"Appareil le plus utilisé: {stats['most_connected_device']}",
                        font=("Roboto", 12)).pack()

        # Filtres
        filters_frame = ctk.CTkFrame(header, fg_color="transparent")
        filters_frame.pack(fill="x", pady=(0, 10))

        # Frame pour les filtres de date
        date_frame = ctk.CTkFrame(filters_frame, fg_color="transparent")
        date_frame.pack(side="left")

        ctk.CTkLabel(date_frame, text="Du:").pack(side="left", padx=(0, 5))
        self.start_date = DateEntry(date_frame, width=12, background='darkblue',
                                  foreground='white', borderwidth=2,
                                  date_pattern='dd/mm/y')
        self.start_date.pack(side="left", padx=5)

        # Sélecteurs d'heure de début
        hour_var = tk.StringVar(value="00")
        minute_var = tk.StringVar(value="00")
        
        self.start_hour = ttk.Spinbox(date_frame, from_=0, to=23, width=3, 
                                    format="%02.0f", textvariable=hour_var)
        self.start_hour.pack(side="left", padx=2)
        
        ctk.CTkLabel(date_frame, text=":").pack(side="left")
        
        self.start_minute = ttk.Spinbox(date_frame, from_=0, to=59, width=3,
                                      format="%02.0f", textvariable=minute_var)
        self.start_minute.pack(side="left", padx=2)

        ctk.CTkLabel(date_frame, text="Au:").pack(side="left", padx=(10, 5))
        self.end_date = DateEntry(date_frame, width=12, background='darkblue',
                                foreground='white', borderwidth=2,
                                date_pattern='dd/mm/y')
        self.end_date.pack(side="left", padx=5)

        # Sélecteurs d'heure de fin
        end_hour_var = tk.StringVar(value="23")
        end_minute_var = tk.StringVar(value="59")
        
        self.end_hour = ttk.Spinbox(date_frame, from_=0, to=23, width=3,
                                  format="%02.0f", textvariable=end_hour_var)
        self.end_hour.pack(side="left", padx=2)
        
        ctk.CTkLabel(date_frame, text=":").pack(side="left")
        
        self.end_minute = ttk.Spinbox(date_frame, from_=0, to=59, width=3,
                                    format="%02.0f", textvariable=end_minute_var)
        self.end_minute.pack(side="left", padx=2)

        # Boutons de filtrage
        buttons_frame = ctk.CTkFrame(filters_frame, fg_color="transparent")
        buttons_frame.pack(side="right")

        ctk.CTkButton(buttons_frame, text="Filtrer", 
                     command=self.apply_date_filter,
                     width=70).pack(side="left", padx=5)

        ctk.CTkButton(buttons_frame, text="Réinitialiser", 
                     command=self.reset_filters,
                     width=90).pack(side="left", padx=5)

    def reset_filters(self):
        # Réinitialiser les dates
        today = datetime.now()
        self.start_date.set_date(today)
        self.end_date.set_date(today)
        
        # Réinitialiser les heures
        self.start_hour.set("00")
        self.start_minute.set("00")
        self.end_hour.set("23")
        self.end_minute.set("59")
        
        # Recharger tout l'historique
        self.load_history()

    def get_datetime_from_widgets(self, date_widget, hour_widget, minute_widget):
        """Convertit les valeurs des widgets en objet datetime"""
        date = date_widget.get_date()
        try:
            hour = int(hour_widget.get())
            minute = int(minute_widget.get())
            return datetime.combine(date, time(hour=min(hour, 23), minute=min(minute, 59)))
        except (ValueError, TypeError):
            # En cas d'erreur, utiliser 00:00 pour le début ou 23:59 pour la fin
            if hour_widget == self.start_hour:
                return datetime.combine(date, time(0, 0))
            else:
                return datetime.combine(date, time(23, 59))

    def apply_date_filter(self):
        start_datetime = self.get_datetime_from_widgets(
            self.start_date, self.start_hour, self.start_minute)
        end_datetime = self.get_datetime_from_widgets(
            self.end_date, self.end_hour, self.end_minute)
        
        self.load_history(start_datetime, end_datetime)

    def create_table(self):
        """Crée le tableau d'historique avec des colonnes fixes"""
        # Créer le frame scrollable pour l'historique
        self.history_frame = ctk.CTkScrollableFrame(self.window)
        self.history_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Créer un frame pour contenir les colonnes
        self.columns_frame = ctk.CTkFrame(self.history_frame, fg_color="transparent")
        self.columns_frame.pack(fill="both", expand=True)

        # Définir les largeurs fixes pour chaque colonne (en pixels)
        self.column_widths = {
            "Date": 100,
            "Heure": 80,
            "Client": 200,
            "Nom du PC": 200
        }

        # En-têtes des colonnes
        headers = list(self.column_widths.keys())
        for i, header in enumerate(headers):
            column_frame = ctk.CTkFrame(self.columns_frame, fg_color="transparent", width=self.column_widths[header])
            column_frame.grid(row=0, column=i, sticky="ew", padx=5)
            column_frame.grid_propagate(False)  # Empêcher le redimensionnement automatique
            
            label = ctk.CTkLabel(column_frame, 
                               text=header,
                               font=("Roboto", 12, "bold"),
                               anchor="center")
            label.pack(fill="both", expand=True)

        # Configurer les poids des colonnes
        for i in range(len(headers)):
            self.columns_frame.grid_columnconfigure(i, weight=1)

    def add_history_entry(self, row_index, log_data):
        """Ajoute une ligne dans le tableau avec des colonnes fixes"""
        values = [
            log_data['date'],
            log_data['time'],
            log_data['client'],
            log_data['device_name']
        ]

        for i, value in enumerate(values):
            # Créer un frame de largeur fixe pour chaque cellule
            cell_frame = ctk.CTkFrame(self.columns_frame, 
                                    fg_color="transparent",
                                    width=self.column_widths[list(self.column_widths.keys())[i]])
            cell_frame.grid(row=row_index, column=i, sticky="ew", padx=5, pady=2)
            cell_frame.grid_propagate(False)  # Empêcher le redimensionnement automatique

            # Ajouter le label dans le frame
            label = ctk.CTkLabel(cell_frame, 
                               text=value,
                               anchor="center")
            label.pack(fill="both", expand=True)

    def load_history(self, start_datetime=None, end_datetime=None):
        """Charge l'historique avec filtrage par date et heure"""
        # Nettoyer la liste existante (sauf les en-têtes)
        for widget in self.columns_frame.grid_slaves():
            if int(widget.grid_info()["row"]) > 0:  # Garder les en-têtes
                widget.destroy()

        # Charger l'historique
        logs = self.connection_logger.get_connection_history()

        if not logs:
            empty_label = ctk.CTkLabel(self.columns_frame, 
                                     text="Aucune connexion trouvée",
                                     text_color="gray")
            empty_label.grid(row=1, column=0, columnspan=4, pady=20)
            return

        # Appliquer les filtres de date et heure si spécifiés
        if start_datetime or end_datetime:
            filtered_logs = []
            for log in logs:
                try:
                    log_datetime = datetime.strptime(
                        f"{log['date']} {log['time']}", 
                        '%d/%m/%Y %H:%M'
                    )
                    
                    if start_datetime and log_datetime < start_datetime:
                        continue
                    if end_datetime and log_datetime > end_datetime:
                        continue
                    filtered_logs.append(log)
                except ValueError:
                    continue
                    
            logs = filtered_logs

        # Trier les logs par date et heure (plus récent en premier)
        logs.sort(key=lambda x: f"{x['date']} {x['time']}", reverse=True)

        if not logs:
            empty_label = ctk.CTkLabel(self.columns_frame, 
                                     text="Aucune connexion trouvée pour cette période",
                                     text_color="gray")
            empty_label.grid(row=1, column=0, columnspan=4, pady=20)
            return

        # Afficher chaque connexion
        for row_index, log in enumerate(logs, start=1):
            self.add_history_entry(row_index, log)

    def export_to_excel(self):
        try:
            logs = self.connection_logger.get_connection_history()
            if not logs:
                messagebox.showwarning("Export", "Aucune donnée à exporter")
                return

            df = pd.DataFrame(logs)
            
            # Réorganiser les colonnes
            df = df[['date', 'time', 'client', 'device_name']]
            
            # Renommer les colonnes
            df.columns = ['Date', 'Heure', 'Client', 'Nom du PC']

            # Demander à l'utilisateur où sauvegarder le fichier
            file_path = filedialog.asksaveasfilename(
                defaultextension='.xlsx',
                filetypes=[("Excel files", "*.xlsx")],
                title="Exporter l'historique"
            )
            
            if file_path:
                df.to_excel(file_path, index=False)
                messagebox.showinfo("Export", "Export réussi !")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'export: {str(e)}")

    def clean_old_entries(self):
        days = 30  # Par défaut, supprimer les entrées de plus de 30 jours
        deleted_count = self.connection_logger.delete_old_entries(days)
        messagebox.showinfo("Nettoyage", 
                          f"{deleted_count} entrées plus anciennes que {days} jours ont été supprimées")
        self.load_history()
        self.update_stats()

    def update_stats(self):
        # Nettoyer les stats existantes
        for widget in self.stats_frame.winfo_children():
            widget.destroy()

        stats = self.connection_logger.get_statistics()
        if not stats:
            return

        # Créer une grille de statistiques
        grid_frame = ctk.CTkFrame(self.stats_frame, fg_color="transparent")
        grid_frame.pack(fill="x", pady=10)

        # Configuration des poids des colonnes
        grid_frame.grid_columnconfigure((0, 1, 2), weight=1)

        # Première ligne
        ctk.CTkLabel(grid_frame, text=f"Total connexions: {stats['total_connections']}", 
                    font=("Roboto", 12)).grid(row=0, column=0, padx=10)
        ctk.CTkLabel(grid_frame, text=f"Clients uniques: {stats['unique_clients']}", 
                    font=("Roboto", 12)).grid(row=0, column=1, padx=10)
        ctk.CTkLabel(grid_frame, text=f"Appareils uniques: {stats['unique_devices']}", 
                    font=("Roboto", 12)).grid(row=0, column=2, padx=10)

        # Deuxième ligne
        ctk.CTkLabel(grid_frame, text=f"Client le plus connecté: {stats['most_connected_client']}", 
                    font=("Roboto", 12)).grid(row=1, column=0, padx=10)
        ctk.CTkLabel(grid_frame, text=f"Appareil le plus utilisé: {stats['most_connected_device']}", 
                    font=("Roboto", 12)).grid(row=1, column=1, columnspan=2, padx=10)

class ModernRustDeskGUI:
    def __init__(self):
        ctk.set_appearance_mode("dark")
        self.style = {
            "padding": 20,
            "button_height": 40,
            "corner_radius": 10
        }
        
        # Initialiser la fenêtre principale
        self.root = ctk.CTk()
        self.root.title("RustDesk - Interface de Connexion")
        self.root.geometry("800x600")
        self.root.minsize(800, 600)
        
        # Vérifier et charger le carnet d'adresses
        try:
            self.address_book = pd.read_excel(r"C:\Windows\address_book.xlsx")
            if self.address_book.empty:
                self.show_error_and_exit("Le carnet d'adresses est vide.")
        except FileNotFoundError:
            self.show_error_and_exit(
                "Carnet d'adresses introuvable !\n\n"
                "Merci de vérifier que le fichier 'address_book.xlsx' "
                "est bien présent dans le dossier C:\\Windows"
            )
        except Exception as e:
            self.show_error_and_exit(f"Erreur lors de la lecture du carnet d'adresses : {str(e)}")

        # Initialiser le logger de connexions
        self.connection_logger = ConnectionLogger()
        
        # Variables d'état
        self.current_devices = []
        self.selected_client = None
        self.client_buttons = {}
        
        # Trouver le chemin de RustDesk
        self.RUSTDESK_PATH = self.get_rustdesk_path()
        if not self.RUSTDESK_PATH:
            self.show_error_and_exit(
                "RustDesk n'est pas installé !\n\n"
                "Veuillez installer RustDesk avant d'utiliser cette application."
            )
        
        # Créer l'interface
        self.create_interface()
        
        # Charger les données
        self.load_data()

    def load_data(self):
        """Charge les données initiales"""
        # Obtenir la liste unique des clients
        clients = self.address_book['Client'].unique()
        
        # Mettre à jour l'interface avec les clients
        self.update_clients_list(clients)
        
        # Mettre à jour le compteur de clients
        self.clients_count.configure(text=f"{len(clients)} clients")

    def show_error_and_exit(self, message):
        """Affiche une fenêtre d'erreur et quitte l'application"""
        root = ctk.CTk()
        root.withdraw()  # Cacher la fenêtre principale vide
        
        # Créer une nouvelle fenêtre pour le message d'erreur
        error_window = ctk.CTkToplevel()
        error_window.title("Erreur")
        error_window.geometry("400x200")
        
        # Centrer la fenêtre
        screen_width = error_window.winfo_screenwidth()
        screen_height = error_window.winfo_screenheight()
        x = (screen_width - 400) // 2
        y = (screen_height - 200) // 2
        error_window.geometry(f"400x200+{x}+{y}")
        
        # Ajouter le message d'erreur
        label = ctk.CTkLabel(error_window, 
                           text=message,
                           wraplength=350,
                           justify="center")
        label.pack(pady=20, padx=20, expand=True)
        
        # Ajouter un bouton pour fermer
        button = ctk.CTkButton(error_window,
                             text="Fermer",
                             command=lambda: self.quit_app(root))
        button.pack(pady=20)
        
        # Empêcher le redimensionnement
        error_window.resizable(False, False)
        
        # Garder la fenêtre au premier plan
        error_window.lift()
        error_window.attributes('-topmost', True)
        
        root.mainloop()
        
    def quit_app(self, root):
        """Quitte proprement l'application"""
        root.quit()
        sys.exit()

    def create_interface(self):
        self.main_container = ctk.CTkFrame(self.root, corner_radius=self.style["corner_radius"])
        self.main_container.pack(fill="both", expand=True, padx=20, pady=20)

        self.create_header()

        columns_container = ctk.CTkFrame(self.main_container, fg_color="transparent")
        columns_container.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.create_left_column(columns_container)

        separator = ctk.CTkFrame(columns_container, width=2, fg_color=("gray75", "gray25"))
        separator.pack(side="left", fill="y", padx=10, pady=10)

        self.create_right_column(columns_container)

        self.create_status_bar()

    def create_header(self):
        header = ctk.CTkFrame(self.main_container, fg_color="transparent")
        header.pack(fill="x", padx=20, pady=(10, 20))

        title = ctk.CTkLabel(header, text="RustDesk Connect", 
                           font=("Roboto", 24, "bold"))
        title.pack(side="left")

        history_btn = ctk.CTkButton(header, text="Historique", 
                                  command=self.show_history,
                                  width=100)
        history_btn.pack(side="right", padx=(0, 10))

        self.time_label = ctk.CTkLabel(header, text="", font=("Roboto", 12))
        self.time_label.pack(side="right", padx=10)

    def create_left_column(self, parent):
        left_frame = ctk.CTkFrame(parent, corner_radius=self.style["corner_radius"])
        left_frame.pack(side="left", fill="both", expand=True)

        header = ctk.CTkFrame(left_frame, fg_color="transparent")
        header.pack(fill="x", padx=10, pady=10)
        
        clients_label = ctk.CTkLabel(header, text="Clients", 
                                   font=("Roboto", 20, "bold"))
        clients_label.pack(side="left", padx=10)

        self.clients_count = ctk.CTkLabel(header, text="0 clients", 
                                        font=("Roboto", 12))
        self.clients_count.pack(side="right", padx=10)

        search_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        search_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.search_var = ctk.StringVar()
        self.search_var.trace('w', self.filter_clients)
        self.search_entry = ctk.CTkEntry(search_frame, 
                                       placeholder_text=" Rechercher un client...",
                                       height=35,
                                       textvariable=self.search_var)
        self.search_entry.pack(fill="x")

        self.clients_list = ctk.CTkScrollableFrame(left_frame)
        self.clients_list.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def create_right_column(self, parent):
        right_frame = ctk.CTkFrame(parent, corner_radius=self.style["corner_radius"])
        right_frame.pack(side="right", fill="both", expand=True)

        devices_header = ctk.CTkFrame(right_frame, fg_color="transparent")
        devices_header.pack(fill="x", padx=10, pady=10)

        self.devices_title = ctk.CTkLabel(devices_header, 
                                        text="Appareils", 
                                        font=("Roboto", 20, "bold"))
        self.devices_title.pack(side="left", padx=10)

        self.devices_count = ctk.CTkLabel(devices_header, 
                                        text="0 appareils", 
                                        font=("Roboto", 12))
        self.devices_count.pack(side="right", padx=10)

        self.devices_frame = ctk.CTkScrollableFrame(right_frame)
        self.devices_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.create_connection_section(right_frame)

    def create_connection_section(self, parent):
        connection_frame = ctk.CTkFrame(parent, fg_color="transparent")
        connection_frame.pack(fill="x", padx=10, pady=10)

        self.password_entry = ctk.CTkEntry(connection_frame, 
                                         placeholder_text="Mot de passe",
                                         show="•",
                                         height=35)
        self.password_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.password_entry.bind('<Return>', lambda event: self.connect())

        self.connect_button = ctk.CTkButton(connection_frame, 
                                          text="Se connecter",
                                          command=self.connect,
                                          height=35)
        self.connect_button.pack(side="right")

    def create_status_bar(self):
        status_frame = ctk.CTkFrame(self.root, height=30, fg_color=("gray85", "gray20"))
        status_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        self.status_label = ctk.CTkLabel(status_frame, text="Prêt", 
                                       font=("Roboto", 12))
        self.status_label.pack(pady=5)

    def update_time(self):
        current_time = datetime.now().strftime("%H:%M:%S")
        self.time_label.configure(text=current_time)
        self.root.after(1000, self.update_time)

    def update_clients_list(self, clients):
        for widget in self.clients_list.winfo_children():
            widget.destroy()
        
        self.client_buttons = {}
        
        for client in clients:
            client_frame = ctk.CTkFrame(self.clients_list, fg_color="transparent")
            client_frame.pack(fill="x", pady=2)
            
            btn = ctk.CTkButton(client_frame, 
                              text=client,
                              command=lambda c=client: self.show_devices(c),
                              fg_color="transparent",
                              text_color=("gray10", "gray90"),
                              hover_color=("gray70", "gray30"),
                              height=self.style["button_height"],
                              anchor="w")
            btn.pack(fill="x")
            self.client_buttons[client] = btn
        
        self.clients_count.configure(text=f"{len(clients)} clients")

    def show_devices(self, client):
        if self.selected_client in self.client_buttons:
            self.client_buttons[self.selected_client].configure(fg_color="transparent")
        self.client_buttons[client].configure(fg_color=("gray75", "gray25"))
        self.selected_client = client

        self.devices_title.configure(text=f"Appareils de {client}")
        
        for widget in self.devices_frame.winfo_children():
            widget.destroy()

        devices = self.address_book[self.address_book['Client'] == client]
        self.current_devices = []

        for _, row in devices.iterrows():
            device_frame = ctk.CTkFrame(self.devices_frame, fg_color="transparent")
            device_frame.pack(fill="x", pady=5)
            
            device_info = f"{row['Nom du PC']} (ID: {row['Identifiant']})"
            self.current_devices.append(row)
            
            device_btn = ctk.CTkButton(device_frame,
                                     text=device_info,
                                     command=lambda d=row: self.select_device(d),
                                     fg_color="transparent",
                                     text_color=("gray10", "gray90"),
                                     hover_color=("gray70", "gray30"),
                                     height=self.style["button_height"],
                                     anchor="w")
            device_btn.pack(fill="x")
        
        self.devices_count.configure(text=f"{len(devices)} appareils")

    def select_device(self, device):
        self.selected_device = device
        self.password_entry.focus()
        self.show_status(f"Appareil sélectionné : {device['Nom du PC']}", "info")

    def connect(self):
        if not self.RUSTDESK_PATH:
            self.show_status("RustDesk n'est pas installé", "error")
            return

        if not hasattr(self, 'selected_device'):
            self.show_status("Veuillez sélectionner un appareil", "error")
            return

        password = self.password_entry.get().strip()
        if not password:
            self.show_status("Le mot de passe ne peut pas être vide", "error")
            return

        # Enregistrer la connexion
        self.connection_logger.log_connection(
            client=self.selected_client,
            device_name=self.selected_device['Nom du PC'],
            device_id=self.selected_device['Identifiant']
        )

        # Lancer RustDesk avec l'ID
        try:
            subprocess.Popen([
                str(self.RUSTDESK_PATH),
                "--connect",
                str(self.selected_device['Identifiant']),
                "--password",
                str(password)
            ])
            
            self.show_status(f"Connexion à {self.selected_device['Nom du PC']}...")
            
            # Programmer l'enregistrement de la déconnexion après un délai
            self.root.after(1000, self.check_connection_status)
        except Exception as e:
            self.show_status(f"Erreur lors de la connexion : {str(e)}", "error")

    def check_connection_status(self):
        # Vérifier si le processus RustDesk est toujours en cours
        rustdesk_running = False
        for proc in psutil.process_iter(['name']):
            try:
                if proc.info['name'].lower() == 'rustdesk.exe':
                    rustdesk_running = True
                    break
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        
        if not rustdesk_running:
            # Si RustDesk n'est plus en cours d'exécution, enregistrer la déconnexion
            # self.connection_logger.log_disconnection(
            #     client=self.selected_client,
            #     device_name=self.selected_device['Nom du PC'],
            #     device_id=self.selected_device['Identifiant']
            # )
            self.show_status("Déconnecté")
        else:
            # Vérifier à nouveau dans 1 seconde
            self.root.after(1000, self.check_connection_status)

    def show_status(self, message, status_type="info"):
        colors = {
            "error": "#FF5252",
            "success": "#4CAF50",
            "info": "gray"
        }
        self.status_label.configure(text=message, text_color=colors.get(status_type, "gray"))

    def get_rustdesk_path(self):
        paths = [
            r"C:\Program Files\RustDesk\rustdesk.exe",
            r"C:\Program Files (x86)\RustDesk\rustdesk.exe"
        ]
        for path in paths:
            if os.path.exists(path):
                return path
        return None

    def filter_clients(self, *args):
        search_term = self.search_var.get().lower()
        filtered_clients = [client for client in self.address_book['Client'].unique() if search_term in client.lower()]
        self.update_clients_list(filtered_clients)

    def show_history(self):
        HistoryWindow(self.root, self.connection_logger)

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ModernRustDeskGUI()
    app.run()