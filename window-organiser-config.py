import pygetwindow as gw
import tkinter as tk
from tkinter import ttk
import json
import os
import logging
import time
import AppOpener

class WindowOrganizer:
    def __init__(self):
        self.setup_logging()
        self.setup_config()
        self.setup_gui()
        self.refresh_all()

    def setup_logging(self):
        """Initialize logging"""
        log_file = 'organizer_config.log'
        
        # Clear existing log file
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write('')  # Clear the file
            
        logging.basicConfig(
            filename=log_file,
            level=logging.INFO,
            format='%(asctime)s - %(message)s',
            encoding='utf-8'
        )
        self.logger = logging.getLogger("OrganizerConfig")
        self.logger.info("=== New Configuration Session Started ===")

    def setup_config(self):
        self.config_file = "window_config.json"
        self.window_configs = self.load_config()

    def setup_gui(self):
        # Main window setup
        self.root = tk.Tk()
        self.root.title("Window Organizer")
        self.root.geometry("1000x700")  # Initial size
        self.root.configure(bg='#1e1e1e')
        self.root.minsize(800, 600)  # Set minimum window size

        # Set custom fonts
        self.title_font = ('Segoe UI', 24, 'bold')
        self.header_font = ('Segoe UI', 14, 'bold')
        self.button_font = ('Segoe UI', 11)
        self.list_font = ('Segoe UI', 11)
        self.status_font = ('Segoe UI', 10)

        # Style configuration
        self.setup_styles()
        
        # Main container
        self.main_frame = ttk.Frame(self.root, padding="30", style='Custom.TFrame')
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        
        # Configure main window grid
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # Configure main frame grid
        self.main_frame.grid_rowconfigure(1, weight=1)  # Lists row expands
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=1)

        # Create content frames
        self.create_title()  # Row 0
        self.create_split_view()  # Row 1
        self.create_buttons()  # Row 2
        self.create_status_bar()  # Row 3

    def setup_styles(self):
        """Configure ttk styles with modern look"""
        self.style = ttk.Style()
        self.style.configure('Custom.TFrame', 
                           background='#1e1e1e')
        
        self.style.configure('Custom.TButton',
                           padding=10,
                           font=self.button_font,
                           background='#007acc',
                           foreground='white')
        
        # Custom colors
        self.colors = {
            'bg_dark': '#1e1e1e',
            'bg_light': '#252526',
            'accent': '#007acc',
            'accent_hover': '#005999',
            'text': '#ffffff',
            'text_secondary': '#cccccc',
            'border': '#333333'
        }

    def create_title(self):
        """Create the title label with enhanced styling"""
        title_frame = tk.Frame(self.main_frame, bg=self.colors['bg_dark'])
        title_frame.grid(row=0, column=0, columnspan=2, pady=(0, 30))

        title_label = tk.Label(title_frame,
                             text="Window Organization Manager",
                             font=self.title_font,
                             bg=self.colors['bg_dark'],
                             fg=self.colors['text'])
        title_label.pack()

        subtitle_label = tk.Label(title_frame,
                                text="Organize and manage your windows effortlessly",
                                font=self.status_font,
                                bg=self.colors['bg_dark'],
                                fg=self.colors['text_secondary'])
        subtitle_label.pack(pady=(5, 0))

    def create_split_view(self):
        """Create the split view with two listboxes"""
        # Container frame for lists with correct background
        lists_container = tk.Frame(self.main_frame, bg=self.colors['bg_dark'])
        lists_container.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(0, 20))
        
        lists_container.grid_columnconfigure(0, weight=1)
        lists_container.grid_columnconfigure(1, weight=1)
        lists_container.grid_rowconfigure(0, weight=1)

        # Left side - Active Windows (using tk.Frame instead of ttk.Frame)
        self.list_frame_left = tk.Frame(lists_container, bg=self.colors['bg_dark'])
        self.list_frame_left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        # Right side - Saved Configurations (using tk.Frame instead of ttk.Frame)
        self.list_frame_right = tk.Frame(lists_container, bg=self.colors['bg_dark'])
        self.list_frame_right.grid(row=0, column=1, sticky="nsew")
        
        # Create listboxes
        self.create_listbox_section(self.list_frame_left, "Active Windows", "window_listbox")
        self.create_listbox_section(self.list_frame_right, "Saved Configurations", "saved_listbox")

    def create_listbox_section(self, parent, title, listbox_name):
        """Create a styled listbox section"""
        # Configure parent frame grid
        parent.grid_rowconfigure(1, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        # Header
        header = tk.Label(parent,
                         text=title,
                         font=self.header_font,
                         bg=self.colors['bg_dark'],
                         fg=self.colors['text'])
        header.grid(row=0, column=0, pady=(0, 10), sticky="w")

        # Listbox container with border
        listbox_frame = tk.Frame(parent,
                               bg=self.colors['border'],
                               padx=1, pady=1)
        listbox_frame.grid(row=1, column=0, sticky="nsew")
        
        # Configure listbox container grid
        listbox_frame.grid_rowconfigure(0, weight=1)
        listbox_frame.grid_columnconfigure(0, weight=1)

        # Listbox
        listbox = tk.Listbox(listbox_frame,
                            font=self.list_font,
                            bg=self.colors['bg_light'],
                            fg=self.colors['text'],
                            selectmode=tk.EXTENDED,
                            activestyle='none',
                            selectbackground=self.colors['accent'],
                            selectforeground='white',
                            borderwidth=0,
                            highlightthickness=0)
        
        # Scrollbar
        scrollbar = tk.Scrollbar(listbox_frame,
                               orient="vertical",
                               command=listbox.yview,
                               width=12,
                               troughcolor=self.colors['bg_light'],
                               bg=self.colors['bg_dark'])
        
        listbox.configure(yscrollcommand=scrollbar.set)
        
        # Place listbox and scrollbar
        listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        setattr(self, listbox_name, listbox)

    def create_buttons(self):
        """Create modern styled buttons"""
        button_frame = tk.Frame(self.main_frame, bg=self.colors['bg_dark'])
        button_frame.grid(row=2, column=0, columnspan=2, sticky="ew")
        
        # Configure button frame grid
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        button_frame.grid_columnconfigure(2, weight=1)
        button_frame.grid_columnconfigure(3, weight=1)
        
        buttons = [
            ("Refresh Windows", self.refresh_windows, '#4CAF50'),
            ("Save Selected", self.save_window, '#2196F3'),
            ("Apply Layouts", self.apply_layouts, '#FF9800'),
            ("Remove Selected", self.remove_saved, '#f44336')
        ]
        
        for i, (text, command, color) in enumerate(buttons):
            btn = tk.Button(button_frame,
                          text=text,
                          command=command,
                          font=self.button_font,
                          bg=color,
                          fg='white',
                          padx=25,
                          pady=12,
                          relief='flat',
                          cursor='hand2',
                          borderwidth=0)
            btn.grid(row=0, column=i, padx=6, sticky="ew")
            
            btn.bind('<Enter>', 
                    lambda e, b=btn, c=color: self.on_button_hover(b, c))
            btn.bind('<Leave>', 
                    lambda e, b=btn, c=color: self.on_button_leave(b, c))

    def on_button_hover(self, button, color):
        """Darken button color on hover"""
        # Convert color to RGB and darken it
        r = int(color[1:3], 16) * 0.8
        g = int(color[3:5], 16) * 0.8
        b = int(color[5:7], 16) * 0.8
        dark_color = f'#{int(r):02x}{int(g):02x}{int(b):02x}'
        button.configure(bg=dark_color)

    def on_button_leave(self, button, color):
        """Restore original button color"""
        button.configure(bg=color)

    def create_status_bar(self):
        """Create a modern status bar"""
        status_frame = tk.Frame(self.main_frame, 
                              bg=self.colors['bg_light'],
                              height=30)
        status_frame.grid(row=3, column=0, columnspan=2, sticky='ew', pady=(20, 0))
        
        self.status_var = tk.StringVar(value="Ready")
        status_bar = tk.Label(status_frame,
                            textvariable=self.status_var,
                            font=self.status_font,
                            bg=self.colors['bg_light'],
                            fg=self.colors['text_secondary'],
                            padx=10,
                            pady=5)
        status_bar.pack(side='left')

    def load_config(self):
        """Load the configuration file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    content = f.read().strip()
                    if content:  # If file is not empty
                        return json.loads(content)
                    return {}
            else:
                # Only create new file if it doesn't exist
                with open(self.config_file, 'w') as f:
                    json.dump({}, f, indent=4)
                return {}
        except json.JSONDecodeError as e:
            self.logger.error(f"Error reading config file: {e}")
            # Backup the broken file instead of overwriting
            if os.path.exists(self.config_file):
                backup_file = f"{self.config_file}.backup"
                os.rename(self.config_file, backup_file)
                self.logger.info(f"Backed up broken config to {backup_file}")
            # Create new empty config
            with open(self.config_file, 'w') as f:
                json.dump({}, f, indent=4)
            return {}

    def save_config(self):
        """Save the configuration file"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.window_configs, f, indent=4)
            self.logger.info(f"Saved config: {self.window_configs}")
        except Exception as e:
            self.logger.error(f"Failed to save config: {e}")

    def get_window_info(self, window):
        """Get window information and clean app name"""
        # Get base app name and clean it
        app_name = window.title.split(" - ")[0].strip()
        
        # Common app name mappings
        app_mappings = {
            "WhatsApp": "whatsapp",
            "Amis": "discord",  # Discord window title
            "Discord": "discord",
            "Google Chrome": "chrome",
            "Mozilla Firefox": "firefox",
            "Microsoft Edge": "edge",
            "Visual Studio Code": "code",
            "SPOTIFY PREMIUM": "spotify",
            "Spotify": "spotify",
            "Steam": "steam",
            "Explorateur de fichiers": "explorer",
            "File Explorer": "explorer",
            "Messenger": "messenger",
            "SteelSeries GG": "steelseries-gg",
            "Mobile connecté": "Mobile connect",  # Updated mapping for Your Phone app
            "Your Phone": "phone",       # English version
            "Phone Link": "phone",       # Alternative name
            # Add more mappings as needed
        }
        
        # Try to match the app name
        clean_app_name = app_mappings.get(app_name, app_name.lower())
        
        self.logger.info(f"Original title: {window.title}")
        self.logger.info(f"Cleaned app name: {clean_app_name}")
        
        return {
            "x": window.left,
            "y": window.top,
            "width": window.width,
            "height": window.height,
            "app_name": clean_app_name,
            "original_title": window.title  # Save original title for matching
        }

    def refresh_windows(self):
        """Refresh the active windows list"""
        self.window_listbox.delete(0, tk.END)
        for window in gw.getAllWindows():
            if window.title and window.title.strip():
                if window.title not in self.window_configs:
                    self.window_listbox.insert(tk.END, window.title)

    def refresh_saved_list(self):
        """Refresh the saved configurations list"""
        self.saved_listbox.delete(0, tk.END)
        for window_title in self.window_configs.keys():
            self.saved_listbox.insert(tk.END, window_title)

    def refresh_all(self):
        """Refresh both lists"""
        self.refresh_windows()
        self.refresh_saved_list()

    def save_window(self):
        """Save selected window configurations"""
        selections = self.window_listbox.curselection()
        if not selections:
            self.status_var.set("Please select one or more windows first.")
            return

        saved_count = 0
        for selection in selections:
            window_title = self.window_listbox.get(selection)
            windows = gw.getWindowsWithTitle(window_title)
            if not windows:
                self.logger.warning(f"Could not find window: {window_title}")
                continue

            window_info = self.get_window_info(windows[0])
            if window_info:
                self.logger.info(f"Saving window info: {window_info}")
                self.window_configs[window_title] = window_info
                saved_count += 1

        self.logger.info(f"Saving {saved_count} windows to config")
        self.save_config()
        self.refresh_all()
        self.status_var.set(f"Saved configuration for {saved_count} windows")

    def remove_saved(self):
        """Remove selected saved configurations"""
        selections = self.saved_listbox.curselection()
        if not selections:
            self.status_var.set("Please select configurations to remove.")
            return

        for selection in reversed(selections):
            window_title = self.saved_listbox.get(selection)
            if window_title in self.window_configs:
                del self.window_configs[window_title]

        self.save_config()
        self.refresh_all()
        self.status_var.set(f"Removed {len(selections)} configurations")

    def apply_layouts(self):
        """Apply all saved window layouts"""
        # Get list of available apps from AppOpener
        available_apps = AppOpener.give_appnames()
        self.logger.info(f"Available apps: {available_apps}")

        for title, config in self.window_configs.items():
            self.logger.info(f"Applying layout for: {title}")
            
            windows = gw.getWindowsWithTitle(title)
            if not windows:
                try:
                    app_name = config["app_name"]
                    self.logger.info(f"Attempting to launch: {app_name}")
                    
                    # Try different methods to launch the app
                    launch_methods = [
                        lambda: AppOpener.open(app_name, match_closest=True),
                        lambda: AppOpener.open(app_name.replace(" ", "-"), match_closest=True),
                        lambda: AppOpener.open(app_name.replace(" ", ""), match_closest=True),
                        lambda: os.system(f"start {app_name}")
                    ]

                    success = False
                    for method in launch_methods:
                        try:
                            method()
                            success = True
                            break
                        except Exception as e:
                            self.logger.warning(f"Launch attempt failed: {e}")
                    
                    if not success:
                        self.logger.error(f"All launch attempts failed for {app_name}")
                        continue

                    # Wait for window to appear
                    for attempt in range(10):
                        time.sleep(1)
                        # Try to match both by title and original title
                        windows = gw.getWindowsWithTitle(title)
                        if not windows and "original_title" in config:
                            windows = gw.getWindowsWithTitle(config["original_title"])
                        if windows:
                            self.logger.info(f"Window appeared after {attempt + 1}s")
                            break
                        self.logger.info(f"Waiting... ({attempt + 1}/10)")
                except Exception as e:
                    self.logger.error(f"Failed to launch {title}: {e}")
                    continue

            if windows:
                window = windows[0]
                try:
                    window.resizeTo(config["width"], config["height"])
                    window.moveTo(config["x"], config["y"])
                    self.logger.info(f"Positioned: {title}")
                except Exception as e:
                    self.logger.error(f"Failed to position {title}: {e}")
            else:
                self.logger.warning(f"Could not find window: {title}")

def main():
    app = WindowOrganizer()
    app.root.mainloop()

if __name__ == "__main__":
    main()
