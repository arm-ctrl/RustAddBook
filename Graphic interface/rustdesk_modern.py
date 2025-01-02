import os
import pandas as pd
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from tkcalendar import DateEntry
from datetime import datetime, time, timedelta
import json
import subprocess
import psutil
import webbrowser
import sys
import signal
import platform

# Dictionnaire des traductions
TRANSLATIONS = {
    'fr': {
        'title': "RustAddBook",
        'history': "Historique",
        'settings': "Paramètres",
        'search_placeholder': "Rechercher un client...",
        'connection_status': "État de la connexion : ",
        'connected': "Connecté",
        'disconnected': "Déconnecté",
        'select_client': "Sélectionner un client",
        'select_device': "Sélectionner un appareil",
        'select_device_of': "Appareils de",
        'connect_button': "Se connecter",
        'password_prompt': "Mot de passe",
        'language': "Langue",
        'french': "Français",
        'english': "Anglais",
        'spanish': "Espagnol",
        'save_settings': "Enregistrer",
        'cancel': "Annuler",
        'error_title': "Erreur",
        'success': "Succès",
        'settings_saved': "Paramètres enregistrés",
        'restart_required': "Redémarrage requis",
        'restart_message': "Veuillez redémarrer l'application pour appliquer les changements",
        'history_title': "Historique des connexions",
        'date': "Date",
        'time': "Heure",
        'client': "Client",
        'device_name': "Nom de l'appareil",
        'device_id': "Identifiant",
        'export': "Exporter",
        'refresh': "Rafraîchir",
        'stats_title': "Statistiques",
        'unique_clients': "Clients",
        'unique_devices': "Appareils",
        'devices': "Appareils",
        'filter_date': "Filtrer par date",
        'filter_time': "Filtrer par heure",
        'reset_filters': "Réinitialiser",
        'no_device_selected': "Aucun appareil sélectionné",
        'connection_error': "Erreur de connexion",
        'address_book_error': "Erreur de carnet d'adresses",
        'ready': "Prêt",
        'device_selected': "Appareil sélectionné :",
        'help': "Aide",
        'most_connected_client': "Client le plus connecté",
        'most_connected_device': "Appareil le plus connecté",
        'total_connections': "Connexions totales",
        'dark_mode': "Mode sombre",
        'devices_count': "{} appareils",
        'no_history_data': "Aucune donnée à exporter dans l'historique",
        'export_success': "L'historique a été exporté avec succès vers :\n{}",
        'export_error': "Erreur lors de l'exportation de l'historique :\n{}"
    },
    'en': {
        'title': "RustAddBook",
        'history': "History",
        'settings': "Settings",
        'search_placeholder': "Search for a client...",
        'connection_status': "Connection status: ",
        'connected': "Connected",
        'disconnected': "Disconnected",
        'select_client': "Select a client",
        'select_device': "Select a device",
        'select_device_of': "Devices of",
        'connect_button': "Connect",
        'password_prompt': "Password",
        'language': "Language",
        'french': "French",
        'english': "English",
        'spanish': "Spanish",
        'save_settings': "Save",
        'cancel': "Cancel",
        'error_title': "Error",
        'success': "Success",
        'settings_saved': "Settings saved",
        'restart_required': "Restart required",
        'restart_message': "Please restart the application to apply changes",
        'history_title': "Connection History",
        'date': "Date",
        'time': "Time",
        'client': "Client",
        'device_name': "Device Name",
        'device_id': "Device ID",
        'export': "Export",
        'refresh': "Refresh",
        'stats_title': "Statistics",
        'unique_clients': "Clients",
        'unique_devices': "Devices",
        'devices': "Devices",
        'filter_date': "Filter by date",
        'filter_time': "Filter by time",
        'reset_filters': "Reset",
        'no_device_selected': "No device selected",
        'connection_error': "Connection error",
        'address_book_error': "Address book error",
        'ready': "Ready",
        'device_selected': "Selected device:",
        'help': "Help",
        'most_connected_client': "Most connected client",
        'most_connected_device': "Most connected device",
        'total_connections': "Total connections",
        'dark_mode': "Dark mode",
        'devices_count': "{} devices",
        'no_history_data': "No data to export in history",
        'export_success': "History has been successfully exported to:\n{}",
        'export_error': "Error exporting history:\n{}"
    },
    'es': {
        'title': "RustAddBook",
        'history': "Historial",
        'settings': "Ajustes",
        'search_placeholder': "Buscar cliente...",
        'connection_status': "Estado de la conexión: ",
        'connected': "Conectado",
        'disconnected': "Desconectado",
        'select_client': "Seleccionar cliente",
        'select_device': "Seleccionar dispositivo",
        'select_device_of': "Dispositivos de",
        'connect_button': "Conectar",
        'password_prompt': "Contraseña",
        'language': "Idioma",
        'french': "Francés",
        'english': "Inglés",
        'spanish': "Español",
        'save_settings': "Guardar",
        'cancel': "Cancelar",
        'error_title': "Error",
        'success': "Éxito",
        'settings_saved': "Ajustes guardados",
        'restart_required': "Reinicio necesario",
        'restart_message': "Por favor, reinicie la aplicación para aplicar los cambios",
        'history_title': "Historial de conexiones",
        'date': "Fecha",
        'time': "Hora",
        'client': "Cliente",
        'device_name': "Nombre del dispositivo",
        'device_id': "Identificador",
        'export': "Exportar",
        'refresh': "Actualizar",
        'stats_title': "Estadísticas",
        'unique_clients': "Clientes",
        'unique_devices': "Dispositivos",
        'devices': "Dispositivos",
        'filter_date': "Filtrar por fecha",
        'filter_time': "Filtrar por hora",
        'reset_filters': "Restablecer",
        'no_device_selected': "Ningún dispositivo seleccionado",
        'connection_error': "Error de conexión",
        'address_book_error': "Error en la libreta de direcciones",
        'ready': "Listo",
        'device_selected': "Dispositivo seleccionado:",
        'help': "Ayuda",
        'most_connected_client': "Cliente más conectado",
        'most_connected_device': "Dispositivo más conectado",
        'total_connections': "Conexiones totales",
        'dark_mode': "Modo oscuro",
        'devices_count': "{} dispositivos",
        'no_history_data': "No hay datos para exportar en el historial",
        'export_success': "El historial se ha exportado correctamente a:\n{}",
        'export_error': "Error al exportar el historial:\n{}"
    }
}

class ConnectionLogger:
    def __init__(self):
        # Initialize log directory and file
        self.log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
        self.log_file = os.path.join(self.log_dir, 'connection_history.json')
        self.ensure_log_file_exists()
        self.migrate_old_logs()

    def ensure_log_file_exists(self):
        # Create logs directory if it doesn't exist
        if not os.path.exists(self.log_dir):
            os.makedirs(self.log_dir)
        
        # Create log file if it doesn't exist
        if not os.path.exists(self.log_file):
            with open(self.log_file, 'w', encoding='utf-8') as f:
                json.dump([], f)

    def migrate_old_logs(self):
        # Convert old entries to new format
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
                    'time': log.get('time_connection', log.get('time', ''))  # Take time_connection or time if it exists
                }
                new_logs.append(new_log)
                modified = True
            
            if modified:
                with open(self.log_file, 'w', encoding='utf-8') as f:
                    json.dump(new_logs, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Erreur lors de la migration : {str(e)}")

    def log_connection(self, client, device_name, device_id):
        # Add new connection to log
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
        # Get connection history with optional date filtering
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
        # Delete old entries
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
        # Calculate connection statistics
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
        # Calculate connection statistics
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
    def __init__(self, parent, connection_logger, current_language):
        # Initialize history window
        self.window = ctk.CTkToplevel(parent)
        self.window.title(TRANSLATIONS[current_language]['history_title'])
        self.window.geometry("800x600")
        self.window.grab_set()
        
        self.connection_logger = connection_logger
        self.current_language = current_language
        
        # Create header with filters
        self.create_header()
        
        # Create connection table
        self.create_table()
        
        # Load data
        self.load_history()

    def create_header(self):
        # Title and action buttons
        title_frame = ctk.CTkFrame(self.window)
        title_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(title_frame, 
                    text=TRANSLATIONS[self.current_language]['history_title'],
                    font=("Roboto", 20, "bold")).pack(side="left")

        # Action buttons
        action_frame = ctk.CTkFrame(title_frame, fg_color="transparent")
        action_frame.pack(side="right")

        export_button = ctk.CTkButton(action_frame, 
                                    text=TRANSLATIONS[self.current_language]['export'],
                                    command=self.export_to_excel)
        export_button.pack(side="left", padx=5)

        refresh_button = ctk.CTkButton(action_frame, 
                                     text=TRANSLATIONS[self.current_language]['refresh'],
                                     command=self.load_history)
        refresh_button.pack(side="left", padx=5)

        # Statistics section
        stats_frame = ctk.CTkFrame(self.window, fg_color=("gray90", "gray20"))
        stats_frame.pack(fill="x", padx=10, pady=5)

        stats = self.connection_logger.get_connection_stats()
        if stats:
            # First line of stats
            stats_line1 = ctk.CTkFrame(stats_frame, fg_color="transparent")
            stats_line1.pack(fill="x", padx=10, pady=5)
            
            ctk.CTkLabel(stats_line1, 
                        text=f"{TRANSLATIONS[self.current_language]['unique_clients']}: {stats['unique_clients']} • " +
                             f"{TRANSLATIONS[self.current_language]['unique_devices']}: {stats['unique_devices']}",
                        font=("Roboto", 12)).pack()

            # Second line of stats
            stats_line2 = ctk.CTkFrame(stats_frame, fg_color="transparent")
            stats_line2.pack(fill="x", padx=10, pady=5)
            
            ctk.CTkLabel(stats_line2,
                        text=f"{TRANSLATIONS[self.current_language]['most_connected_client']}: {stats['most_connected_client']} • " +
                             f"{TRANSLATIONS[self.current_language]['most_connected_device']}: {stats['most_connected_device']}",
                        font=("Roboto", 12)).pack()

        # Filters section
        filters_frame = ctk.CTkFrame(self.window)
        filters_frame.pack(fill="x", padx=10, pady=5)

        # Start date filter
        start_date_frame = ctk.CTkFrame(filters_frame, fg_color="transparent")
        start_date_frame.pack(side="left", padx=10)

        ctk.CTkLabel(start_date_frame, 
                    text="Du").pack(side="left", padx=5)

        self.start_date = DateEntry(start_date_frame, width=12, background='darkblue',
                                  foreground='white', borderwidth=2,
                                  date_pattern='dd/mm/y')
        self.start_date.pack(side="left", padx=5)

        # Start time filter
        start_time_frame = ctk.CTkFrame(start_date_frame, fg_color="transparent")
        start_time_frame.pack(side="left", padx=5)

        self.start_hour = ttk.Spinbox(start_time_frame, from_=0, to=23, width=3, 
                                    format="%02.0f")
        self.start_hour.pack(side="left", padx=2)
        
        ctk.CTkLabel(start_time_frame, text=":").pack(side="left")
        
        self.start_minute = ttk.Spinbox(start_time_frame, from_=0, to=59, width=3,
                                      format="%02.0f")
        self.start_minute.pack(side="left", padx=2)

        # End date filter
        end_date_frame = ctk.CTkFrame(filters_frame, fg_color="transparent")
        end_date_frame.pack(side="left", padx=10)

        ctk.CTkLabel(end_date_frame, 
                    text="Au").pack(side="left", padx=5)

        self.end_date = DateEntry(end_date_frame, width=12, background='darkblue',
                                foreground='white', borderwidth=2,
                                date_pattern='dd/mm/y')
        self.end_date.pack(side="left", padx=5)

        # End time filter
        end_time_frame = ctk.CTkFrame(end_date_frame, fg_color="transparent")
        end_time_frame.pack(side="left", padx=5)

        self.end_hour = ttk.Spinbox(end_time_frame, from_=0, to=23, width=3, 
                                  format="%02.0f")
        self.end_hour.pack(side="left", padx=2)
        
        ctk.CTkLabel(end_time_frame, text=":").pack(side="left")
        
        self.end_minute = ttk.Spinbox(end_time_frame, from_=0, to=59, width=3,
                                    format="%02.0f")
        self.end_minute.pack(side="left", padx=2)

        # Filter buttons
        buttons_frame = ctk.CTkFrame(filters_frame, fg_color="transparent")
        buttons_frame.pack(side="right", padx=10)

        apply_button = ctk.CTkButton(buttons_frame,
                                   text="OK",
                                   command=self.apply_date_filter,
                                   width=50)
        apply_button.pack(side="left", padx=5)

        reset_button = ctk.CTkButton(buttons_frame,
                                   text=TRANSLATIONS[self.current_language]['reset_filters'],
                                   command=self.reset_filters)
        reset_button.pack(side="left", padx=5)

    def create_table(self):
        # Create scrollable frame for connection history
        self.history_frame = ctk.CTkScrollableFrame(self.window)
        self.history_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Configure grid for the scrollable frame
        self.history_frame.grid_columnconfigure((0,1,2,3,4), weight=1)

        # Create table headers
        headers = ['date', 'time', 'client', 'device_name', 'device_id']
        for i, header in enumerate(headers):
            label = ctk.CTkLabel(self.history_frame,
                            text=TRANSLATIONS[self.current_language][header],
                            font=("Roboto", 12, "bold"))
            label.grid(row=0, column=i, sticky="ew", padx=5)

    def load_history(self, start_datetime=None, end_datetime=None):
        # Clear existing entries
        for widget in self.history_frame.winfo_children():
            if int(widget.grid_info()['row']) > 0:  # Preserve headers
                widget.destroy()

        # Get connection history
        history = self.connection_logger.get_connection_history()
        
        # Sort by timestamp (newest first)
        history.sort(key=lambda x: x['timestamp'], reverse=True)

        # Filter by date range if specified
        if start_datetime or end_datetime:
            filtered_history = []
            for entry in history:
                entry_datetime = datetime.strptime(entry['timestamp'], '%Y-%m-%dT%H:%M:%S.%f')
                if start_datetime and entry_datetime < start_datetime:
                    continue
                if end_datetime and entry_datetime > end_datetime:
                    continue
                filtered_history.append(entry)
            history = filtered_history

        # Add entries to table
        for row_idx, entry in enumerate(history, start=1):
            timestamp = datetime.strptime(entry['timestamp'], '%Y-%m-%dT%H:%M:%S.%f')
            
            # Create labels with grid
            date_label = ctk.CTkLabel(self.history_frame, text=timestamp.strftime("%Y-%m-%d"))
            date_label.grid(row=row_idx, column=0, sticky="ew", padx=5, pady=2)
            
            time_label = ctk.CTkLabel(self.history_frame, text=timestamp.strftime("%H:%M:%S"))
            time_label.grid(row=row_idx, column=1, sticky="ew", padx=5, pady=2)
            
            client_label = ctk.CTkLabel(self.history_frame, text=entry['client'])
            client_label.grid(row=row_idx, column=2, sticky="ew", padx=5, pady=2)
            
            device_name_label = ctk.CTkLabel(self.history_frame, text=entry['device_name'])
            device_name_label.grid(row=row_idx, column=3, sticky="ew", padx=5, pady=2)
            
            device_id_label = ctk.CTkLabel(self.history_frame, text=entry['device_id'])
            device_id_label.grid(row=row_idx, column=4, sticky="ew", padx=5, pady=2)

    def reset_filters(self):
        # Reset filters
        current_date = datetime.now()
        self.start_date.set_date(current_date)
        self.end_date.set_date(current_date)
        self.start_hour.set("00")
        self.start_minute.set("00")
        self.end_hour.set("23")
        self.end_minute.set("59")
        
        # Reload history
        self.load_history()

    def apply_date_filter(self):
        # Apply date filter
        start_datetime = self.get_datetime_from_widgets(
            self.start_date, self.start_hour, self.start_minute)
        end_datetime = self.get_datetime_from_widgets(
            self.end_date, self.end_hour, self.end_minute, is_end=True)
        
        # Reload history with filter
        self.load_history(start_datetime, end_datetime)

    def get_datetime_from_widgets(self, date_widget, hour_widget, minute_widget, is_end=False):
        # Get datetime from widgets
        date = date_widget.get_date()
        try:
            hour = int(hour_widget.get())
            minute = int(minute_widget.get())
            return datetime.combine(date, time(hour=min(hour, 23), minute=min(minute, 59)))
        except (ValueError, TypeError):
            # If error, use 00:00 for start or 23:59 for end
            if not is_end:
                return datetime.combine(date, time(0, 0))
            else:
                return datetime.combine(date, time(23, 59))

    def export_to_excel(self):
        # Export to Excel
        try:
            history = self.connection_logger.get_connection_history()
            if not history:
                messagebox.showwarning(
                    TRANSLATIONS[self.current_language]['error_title'],
                    TRANSLATIONS[self.current_language]['no_history_data']
                )
                return

            df = pd.DataFrame(history)
            
            # Reorganize columns
            df = df[['date', 'time', 'client', 'device_name']]
            
            # Rename columns
            df.columns = [
                TRANSLATIONS[self.current_language]['date'],
                TRANSLATIONS[self.current_language]['time'],
                TRANSLATIONS[self.current_language]['client'],
                TRANSLATIONS[self.current_language]['device_name']
            ]

            # Ask user for file path
            file_path = filedialog.asksaveasfilename(
                defaultextension='.xlsx',
                filetypes=[("Excel files", "*.xlsx")],
                title=TRANSLATIONS[self.current_language]['export']
            )
            
            if file_path:
                df.to_excel(file_path, index=False)
                messagebox.showinfo(
                    TRANSLATIONS[self.current_language]['success'],
                    TRANSLATIONS[self.current_language]['export_success'].format(file_path)
                )
        except Exception as e:
            messagebox.showerror(
                TRANSLATIONS[self.current_language]['error_title'],
                TRANSLATIONS[self.current_language]['export_error'].format(str(e))
            )

class SettingsWindow:
    def __init__(self, parent, current_language, callback):
        # Initialize settings window
        self.window = ctk.CTkToplevel(parent)
        self.window.title(TRANSLATIONS[current_language]['settings'])
        self.window.geometry("400x300")
        self.window.grab_set()
        
        self.current_language = current_language
        self.callback = callback
        
        # Language section
        language_label = ctk.CTkLabel(self.window, text=TRANSLATIONS[current_language]['language'])
        language_label.pack(pady=10)
        
        self.language_var = tk.StringVar(value=current_language)
        
        # Radio buttons frame
        radio_frame = ctk.CTkFrame(self.window)
        radio_frame.pack(pady=10)
        
        fr_radio = ctk.CTkRadioButton(radio_frame, text=TRANSLATIONS[current_language]['french'],
                                    variable=self.language_var, value='fr')
        fr_radio.pack(side=tk.LEFT, padx=10)
        
        en_radio = ctk.CTkRadioButton(radio_frame, text=TRANSLATIONS[current_language]['english'],
                                    variable=self.language_var, value='en')
        en_radio.pack(side=tk.LEFT, padx=10)

        es_radio = ctk.CTkRadioButton(radio_frame, text=TRANSLATIONS[current_language]['spanish'],
                                    variable=self.language_var, value='es')
        es_radio.pack(side=tk.LEFT, padx=10)
        
        # Dark mode toggle
        self.dark_mode_var = tk.BooleanVar(value=True)
        dark_mode_checkbox = ctk.CTkCheckBox(self.window, 
                                           text=TRANSLATIONS[current_language]['dark_mode'],
                                           variable=self.dark_mode_var, 
                                           command=self.toggle_dark_mode)
        dark_mode_checkbox.pack(pady=10)
        
        # Action buttons
        button_frame = ctk.CTkFrame(self.window)
        button_frame.pack(pady=20)
        
        save_button = ctk.CTkButton(button_frame, 
                                  text=TRANSLATIONS[current_language]['save_settings'],
                                  command=self.save_settings)
        save_button.pack(side=tk.LEFT, padx=10)
        
        cancel_button = ctk.CTkButton(button_frame, 
                                    text=TRANSLATIONS[self.current_language]['cancel'],
                                    command=self.window.destroy)
        cancel_button.pack(side=tk.LEFT, padx=10)

    def toggle_dark_mode(self):
        # Toggle dark mode
        if self.dark_mode_var.get():
            ctk.set_appearance_mode("dark")
        else:
            ctk.set_appearance_mode("light")

    def save_settings(self):
        # Save settings
        new_language = self.language_var.get()
        if new_language != self.current_language:
            self.callback(new_language)
            messagebox.showinfo(
                TRANSLATIONS[self.current_language]['restart_required'],
                TRANSLATIONS[self.current_language]['restart_message']
            )
        self.window.destroy()

class ModernRustDeskGUI:
    def __init__(self):
        # Set dark mode by default
        ctk.set_appearance_mode("dark")
        self.style = {
            "padding": 20,
            "button_height": 40,
            "corner_radius": 10
        }
        
        # Default language
        self.current_language = 'fr'
        self.load_language_preference()
        
        # Initialize main window
        self.root = ctk.CTk()
        self.root.title(TRANSLATIONS[self.current_language]['title'])
        self.root.geometry("800x600")
        self.root.minsize(800, 600)
        
        # Load address book
        try:
            address_book_path = self.get_address_book_path()
            if not address_book_path:
                system = platform.system().lower()
                if system == "windows":
                    error_msg = "Address book not found!\nPlease make sure the file exists at: C:\\Windows\\address_book.xlsx"
                elif system == "darwin":
                    error_msg = "Address book not found!\nPlease place the file in one of these locations:\n" \
                              "- /Applications/RustAddBook/address_book.xlsx\n" \
                              "- ~/Library/Application Support/RustAddBook/address_book.xlsx"
                elif system == "linux":
                    error_msg = "Address book not found!\nPlease place the file in one of these locations:\n" \
                              "- ~/rustaddbook/address_book.xlsx\n" \
                              "- ~/.rustaddbook/address_book.xlsx\n" \
                              "\nNote: ~ represents your home directory"
                else:
                    error_msg = "Address book not found and unsupported operating system!"
                self.show_error_and_exit(error_msg)
            self.address_book = pd.read_excel(address_book_path)
            if self.address_book.empty:
                self.show_error_and_exit(TRANSLATIONS[self.current_language]['address_book_error'])
        except FileNotFoundError:
            self.show_error_and_exit(TRANSLATIONS[self.current_language]['address_book_missing'])
        except Exception as e:
            self.show_error_and_exit(f"{TRANSLATIONS[self.current_language]['error_title']}: {str(e)}")

        # Initialize connection logger
        self.connection_logger = ConnectionLogger()
        
        # State variables
        self.current_devices = []
        self.selected_client = None
        self.client_buttons = {}
        
        # Find RustDesk path
        self.RUSTDESK_PATH = self.get_rustdesk_path()
        if not self.RUSTDESK_PATH:
            self.show_error_and_exit(
                "RustDesk is not installed!\n\n"
                "Please install RustDesk before using this application."
            )
        
        # Create interface
        self.create_interface()
        
        # Load initial data
        self.load_data()

    def load_language_preference(self):
        # Load language preference
        try:
            config_dir = os.path.join(os.getenv('APPDATA'), 'RustDeskInterface')
            config_file = os.path.join(config_dir, 'config.json')
            
            if os.path.exists(config_file):
                with open(config_file, 'r') as f:
                    config = json.load(f)
                    self.current_language = config.get('language', 'fr')
        except Exception:
            self.current_language = 'fr'

    def save_language_preference(self, language):
        # Save language preference
        try:
            config_dir = os.path.join(os.getenv('APPDATA'), 'RustDeskInterface')
            if not os.path.exists(config_dir):
                os.makedirs(config_dir)
            
            config_file = os.path.join(config_dir, 'config.json')
            
            config = {'language': language}
            with open(config_file, 'w') as f:
                json.dump(config, f)
        except Exception as e:
            print(f"Erreur lors de la sauvegarde des préférences : {str(e)}")

    def get_address_book_path(self):
        """Get the path to address_book.xlsx based on the operating system."""
        system = platform.system().lower()
        
        # Define paths for different operating systems
        if system == "windows":
            paths = [r"C:\Windows\address_book.xlsx"]
        elif system == "darwin":  # macOS
            paths = [
                "/Applications/RustAddBook/address_book.xlsx",
                os.path.expanduser("~/Library/Application Support/RustAddBook/address_book.xlsx")
            ]
        elif system == "linux":
            paths = [
                os.path.expanduser("~/rustaddbook/address_book.xlsx"),
                os.path.expanduser("~/.rustaddbook/address_book.xlsx")
            ]
        else:
            self.show_error_and_exit(f"Unsupported operating system: {system}")
            return None

        # Check each path
        for path in paths:
            if os.path.exists(path):
                return path
                
        # Create error message based on OS
        if system == "windows":
            error_msg = "Address book not found!\nPlease make sure the file exists at: C:\\Windows\\address_book.xlsx"
        elif system == "darwin":
            error_msg = "Address book not found!\nPlease place the file in one of these locations:\n" \
                      "- /Applications/RustAddBook/address_book.xlsx\n" \
                      "- ~/Library/Application Support/RustAddBook/address_book.xlsx"
        elif system == "linux":
            error_msg = "Address book (address_book.xlsx) not found!\n\n" \
                      "Please place the file in one of these locations:\n" \
                      "1. ~/rustaddbook/address_book.xlsx\n" \
                      "   (Example: /home/*user*/rustaddbook/address_book.xlsx)\n\n" \
                      "2. ~/.rustaddbook/address_book.xlsx\n" \
                      "   (Example: /home/*user*/.rustaddbook/address_book.xlsx)\n\n" \
                      "Note: Replace 'user' with your username"
        
        self.show_error_and_exit(error_msg)
        return None

    def create_interface(self):
        # Create main container
        self.main_container = ctk.CTkFrame(self.root, corner_radius=self.style["corner_radius"])
        self.main_container.pack(fill="both", expand=True, padx=20, pady=20)

        # Create header
        self.create_header()

        # Create columns container
        columns_container = ctk.CTkFrame(self.main_container, fg_color="transparent")
        columns_container.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Create left column
        self.create_left_column(columns_container)

        # Create separator
        separator = ctk.CTkFrame(columns_container, width=2, fg_color=("gray75", "gray25"))
        separator.pack(side="left", fill="y", padx=10, pady=10)

        # Create right column
        self.create_right_column(columns_container)

        # Create status bar
        self.create_status_bar()

    def create_header(self):
        # Create header
        header = ctk.CTkFrame(self.main_container, fg_color="transparent")
        header.pack(fill="x", padx=20, pady=(10, 20))

        # Create title
        title = ctk.CTkLabel(header, text=TRANSLATIONS[self.current_language]['title'], 
                           font=("Roboto", 24, "bold"))
        title.pack(side="left")

        # Create action buttons
        buttons_frame = ctk.CTkFrame(header, fg_color="transparent")
        buttons_frame.pack(side="right")

        history_btn = ctk.CTkButton(buttons_frame, text=TRANSLATIONS[self.current_language]['history'], 
                                  command=self.show_history,
                                  width=100)
        history_btn.pack(side=tk.LEFT, padx=(0, 5))

        settings_btn = ctk.CTkButton(buttons_frame, text="⚙", 
                                  command=self.show_settings,
                                  width=35,
                                  font=("Segoe UI", 16))
        settings_btn.pack(side=tk.LEFT, padx=(0, 5))

        help_btn = ctk.CTkButton(buttons_frame, text="?",
                               command=self.show_help,
                               width=35,
                               font=("Segoe UI", 16))
        help_btn.pack(side=tk.LEFT, padx=(0, 10))

        # Create time label
        self.time_label = ctk.CTkLabel(header, text="", font=("Roboto", 12))
        self.time_label.pack(side="right", padx=10)

    def create_left_column(self, parent):
        # Create left column
        left_frame = ctk.CTkFrame(parent, corner_radius=self.style["corner_radius"])
        left_frame.pack(side="left", fill="both", expand=True)

        # Create header
        header = ctk.CTkFrame(left_frame, fg_color="transparent")
        header.pack(fill="x", padx=10, pady=10)
        
        # Create clients label
        clients_label = ctk.CTkLabel(header, text=TRANSLATIONS[self.current_language]['select_client'], 
                                   font=("Roboto", 20, "bold"))
        clients_label.pack(side="left", padx=10)

        # Create clients count label
        self.clients_count = ctk.CTkLabel(header, text="0 clients", 
                                        font=("Roboto", 12))
        self.clients_count.pack(side="right", padx=10)

        # Create search frame
        search_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        search_frame.pack(fill="x", padx=10, pady=(0, 10))

        # Create search entry
        self.search_var = ctk.StringVar()
        self.search_var.trace('w', self.filter_clients)
        self.search_entry = ctk.CTkEntry(search_frame, 
                                       placeholder_text=TRANSLATIONS[self.current_language]['search_placeholder'],
                                       height=35,
                                       textvariable=self.search_var)
        self.search_entry.pack(fill="x")

        # Create clients list
        self.clients_list = ctk.CTkScrollableFrame(left_frame)
        self.clients_list.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def create_right_column(self, parent):
        # Create right column
        right_frame = ctk.CTkFrame(parent, corner_radius=self.style["corner_radius"])
        right_frame.pack(side="right", fill="both", expand=True)

        # Create devices header
        devices_header = ctk.CTkFrame(right_frame, fg_color="transparent")
        devices_header.pack(fill="x", padx=10, pady=10)

        # Create devices title label
        self.devices_title = ctk.CTkLabel(devices_header, 
                                        text=TRANSLATIONS[self.current_language]['select_device'], 
                                        font=("Roboto", 20, "bold"))
        self.devices_title.pack(side="left", padx=10)

        # Create devices count label
        self.devices_count = ctk.CTkLabel(devices_header, 
                                        text=TRANSLATIONS[self.current_language]['devices_count'].format(0),
                                        font=("Roboto", 12))
        self.devices_count.pack(side="right", padx=10)

        # Create devices frame
        self.devices_frame = ctk.CTkScrollableFrame(right_frame)
        self.devices_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Create connection section
        self.create_connection_section(right_frame)

    def create_connection_section(self, parent):
        # Create connection section
        connection_frame = ctk.CTkFrame(parent, fg_color="transparent")
        connection_frame.pack(fill="x", padx=10, pady=10)

        # Create password entry
        self.password_entry = ctk.CTkEntry(connection_frame, 
                                         placeholder_text=TRANSLATIONS[self.current_language]['password_prompt'],
                                         show="•",
                                         height=35)
        self.password_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.password_entry.bind('<Return>', lambda event: self.connect())

        # Create connect button
        self.connect_button = ctk.CTkButton(connection_frame, 
                                          text=TRANSLATIONS[self.current_language]['connect_button'],
                                          command=self.connect,
                                          height=35)
        self.connect_button.pack(side="right")

    def create_status_bar(self):
        # Create status bar
        status_frame = ctk.CTkFrame(self.root, height=30, fg_color=("gray85", "gray20"))
        status_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        # Create status label
        self.status_label = ctk.CTkLabel(status_frame, text=TRANSLATIONS[self.current_language]['ready'], 
                                       font=("Roboto", 12))
        self.status_label.pack(pady=5)

    def update_time(self):
        # Update time label
        current_time = datetime.now().strftime("%H:%M:%S")
        self.time_label.configure(text=current_time)
        self.root.after(1000, self.update_time)

    def update_clients_list(self, clients):
        # Update clients list
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
        # Show devices for client
        if self.selected_client in self.client_buttons:
            self.client_buttons[self.selected_client].configure(fg_color="transparent")
        self.client_buttons[client].configure(fg_color=("gray75", "gray25"))
        self.selected_client = client

        self.devices_title.configure(text=f"{TRANSLATIONS[self.current_language]['select_device_of']} {client}")
        
        for widget in self.devices_frame.winfo_children():
            widget.destroy()

        devices = self.address_book[self.address_book['Client'] == client]
        self.current_devices = []

        for _, row in devices.iterrows():
            device_frame = ctk.CTkFrame(self.devices_frame, fg_color="transparent")
            device_frame.pack(fill="x", pady=5)
            
            device_info = f"{row['Hostname']} (ID: {row['Rustdesk_ID']})"
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
        
        self.devices_count.configure(text=TRANSLATIONS[self.current_language]['devices_count'].format(len(devices)))

    def select_device(self, device):
        # Select device
        self.selected_device = device
        self.password_entry.focus()
        self.show_status(f"{TRANSLATIONS[self.current_language]['device_selected']} {device['Hostname']}", "info")

    def connect(self):
        # Connect to device
        if not self.RUSTDESK_PATH:
            self.show_status(TRANSLATIONS[self.current_language]['connection_error'], "error")
            return

        if not hasattr(self, 'selected_device'):
            self.show_status(TRANSLATIONS[self.current_language]['no_device_selected'], "error")
            return

        password = self.password_entry.get().strip()
        if not password:
            self.show_status(TRANSLATIONS[self.current_language]['connection_error'], "error")
            return

        # Log connection
        self.connection_logger.log_connection(
            client=self.selected_client,
            device_name=self.selected_device['Hostname'],
            device_id=self.selected_device['Rustdesk_ID']
        )

        # Launch RustDesk with ID
        try:
            subprocess.Popen([
                str(self.RUSTDESK_PATH),
                "--connect",
                str(self.selected_device['Rustdesk_ID']),
                "--password",
                str(password)
            ])
            
            self.show_status(f"{TRANSLATIONS[self.current_language]['connection_status']} {self.selected_device['Hostname']}...")
            
            # Schedule disconnection logging after delay
            self.root.after(1000, self.check_connection_status)
        except Exception as e:
            self.show_status(f"{TRANSLATIONS[self.current_language]['connection_error']}: {str(e)}", "error")

    def check_connection_status(self):
        # Check connection status
        # Check if RustDesk process is still running
        rustdesk_running = False
        for proc in psutil.process_iter(['name']):
            try:
                if proc.info['name'].lower() == 'rustdesk.exe':
                    rustdesk_running = True
                    break
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        
        if not rustdesk_running:
            # If RustDesk is not running, log disconnection
            # self.connection_logger.log_disconnection(
            #     client=self.selected_client,
            #     device_name=self.selected_device['Hostname'],
            #     device_id=self.selected_device['Rustdesk_ID']
            # )
            self.show_status(TRANSLATIONS[self.current_language]['disconnected'])
        else:
            # If RustDesk is still running, check again after delay
            self.root.after(1000, self.check_connection_status)

    def show_status(self, message, status_type="info"):
        # Show status message
        colors = {
            "error": "#FF5252",
            "success": "#4CAF50",
            "info": "gray"
        }
        self.status_label.configure(text=message, text_color=colors.get(status_type, "gray"))

    def show_error_and_exit(self, message):
        """Show error message and exit the application."""
        messagebox.showerror("Error", message)
        if hasattr(self, 'root'):
            self.root.destroy()
        sys.exit(1)

    def get_rustdesk_path(self):
        # Get system platform
        system = platform.system().lower()
        
        # Define paths for different operating systems
        if system == "windows":
            paths = [
                r"C:\Program Files\RustDesk\rustdesk.exe",
                r"C:\Program Files (x86)\RustDesk\rustdesk.exe"
            ]
        elif system == "darwin":  # macOS
            paths = [
                "/Applications/RustDesk.app/Contents/MacOS/RustDesk",
                os.path.expanduser("~/Applications/RustDesk.app/Contents/MacOS/RustDesk")
            ]
        elif system == "linux":
            paths = [
                "/usr/bin/rustdesk",
                "/usr/local/bin/rustdesk",
                os.path.expanduser("~/.local/bin/rustdesk")
            ]
        else:
            self.show_status(f"Unsupported operating system: {system}", "error")
            return None

        # Check each path
        for path in paths:
            if os.path.exists(path):
                return path
                
        self.show_status(f"RustDesk was not found on your system ({system})", "error")
        return None

    def filter_clients(self, *args):
        # Filter clients
        search_term = self.search_var.get().lower()
        filtered_clients = [client for client in self.address_book['Client'].unique() if search_term in client.lower()]
        self.update_clients_list(filtered_clients)

    def show_history(self):
        # Show history window
        HistoryWindow(self.root, self.connection_logger, self.current_language)

    def show_settings(self):
        # Show settings window
        SettingsWindow(self.root, self.current_language, self.change_language)

    def change_language(self, new_language):
        # Change language
        self.current_language = new_language
        self.save_language_preference(new_language)
        
        # Update interface texts
        self.root.title(TRANSLATIONS[new_language]['title'])
        self.devices_title.configure(text=TRANSLATIONS[new_language]['select_device'])
        
        # Update devices count with current number
        current_devices = len(self.current_devices) if hasattr(self, 'current_devices') else 0
        self.devices_count.configure(text=TRANSLATIONS[new_language]['devices_count'].format(current_devices))
        
        messagebox.showinfo(
            TRANSLATIONS[new_language]['settings_saved'],
            TRANSLATIONS[new_language]['restart_message']
        )

    def load_data(self):
        # Load initial data
        # Get unique clients
        clients = self.address_book['Client'].unique()
        
        # Update interface with clients
        self.update_clients_list(clients)
        
        # Update clients count
        stats = self.connection_logger.get_statistics()
        total_clients = len(clients)
        total_devices = len(self.address_book)
        self.clients_count.configure(
            text=f"{TRANSLATIONS[self.current_language]['unique_clients']}: {total_clients} • "
                 f"{TRANSLATIONS[self.current_language]['devices']}: {total_devices}"
        )

    def show_help(self):
        # Show help
        webbrowser.open('https://github.com/arm-ctrl/RustAddBook')

    def run(self):
        # Run application
        self.root.mainloop()

def handle_interrupt(signal, frame):
    print("\nClosing application...")
    sys.exit(0)

if __name__ == "__main__":
    try:
        # Register interrupt handler
        import signal
        signal.signal(signal.SIGINT, handle_interrupt)
        
        app = ModernRustDeskGUI()
        app.run()
    except KeyboardInterrupt:
        print("\nClosing application...")
        sys.exit(0)
    except Exception as e:
        print(f"Error starting application: {str(e)}")
        sys.exit(1)