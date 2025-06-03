import pygetwindow as gw
import json
import time
import AppOpener
import logging
from pathlib import Path
from win10toast import ToastNotifier

class WindowOrganiser:
    TOAST_TITLE = "Window Organizer"
    def __init__(self):
        self.setup_logging()
        self.config_file = "window_config.json"
        self.load_config()
        self.toaster = ToastNotifier()
        
        # Show toast before waiting
        self.toaster.show_toast(
            self.TOAST_TITLE,
            "Waiting 20 seconds before starting...",
            duration=5,
            threaded=False,
            icon_path="window-organiser-icon.ico"  # Use double backslashes or raw string
        )
        # Initial delay to allow system startup
        self.logger.info("Waiting 20 seconds before starting...")
        time.sleep(20)
        self.logger.info("Starting organization process...")

    def setup_logging(self):
        """Initialize logging"""
        log_file = 'organizer_process.log'
        
        # Clear existing log file
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write('')  # Clear the file
            
        logging.basicConfig(
            filename=log_file,
            level=logging.INFO,
            format='%(asctime)s - %(message)s',
            encoding='utf-8'
        )
        self.logger = logging.getLogger("OrganizerProcess")
        self.logger.info("=== New Organization Process Started ===")

    def load_config(self):
        """Load window configurations from JSON"""
        try:
            if Path(self.config_file).exists():
                with open(self.config_file, 'r') as f:
                    content = f.read().strip()
                    self.window_configs = json.loads(content) if content else {}
            else:
                self.window_configs = {}
        except Exception as e:
            self.logger.error(f"Error loading config: {e}")
            self.window_configs = {}

    def launch_and_wait_for_window(self, title, app_name, original_title=None, max_attempts=10, delay=1):
        """Try to launch an app and wait for its window to appear."""
        launch_methods = [
            lambda: AppOpener.open(app_name, match_closest=True),
            lambda: AppOpener.open(app_name.replace(" ", "-"), match_closest=True),
            lambda: AppOpener.open(app_name.replace(" ", ""), match_closest=True)
        ]
        for method in launch_methods:
            try:
                method()
                break
            except Exception as e:
                self.logger.warning(f"Launch attempt failed: {e}")
                continue

        for attempt in range(max_attempts):
            time.sleep(delay)
            windows = gw.getWindowsWithTitle(title)
            if not windows and original_title:
                windows = gw.getWindowsWithTitle(original_title)
            if windows:
                self.logger.info(f"Window appeared after {attempt + 1}s")
                return windows[0]
            self.logger.info(f"Waiting... ({attempt + 1}/{max_attempts})")
        return None

    def position_window(self, window, config, title):
        try:
            window.resizeTo(config["width"], config["height"])
            window.moveTo(config["x"], config["y"])
            self.logger.info(f"Positioned: {title}")
            return True
        except Exception as e:
            self.logger.error(f"Failed to position {title}: {e}")
            return False

    def organize_windows(self):
        """Open and organize all windows according to saved configuration"""
        # Show start notification (non-threaded to ensure it shows)
        self.toaster.show_toast(
            self.TOAST_TITLE,
            "Starting to organize windows...",
            duration=3,
            threaded=False
        )
        
        self.logger.info("Starting window organization")
        success_count = 0
        total_windows = len(self.window_configs)
        
        for title, config in self.window_configs.items():
            self.logger.info(f"Processing: {title}")
            
            # Check if window exists
            windows = gw.getWindowsWithTitle(title)
            
            # If window doesn't exist, try to launch it
            if not windows:
                window = self.launch_and_wait_for_window(
                    title, config["app_name"], config.get("original_title")
                )
            else:
                window = windows[0]

            # Position window if found
            if window and self.position_window(window, config, title):
                success_count += 1
            else:
                self.logger.warning(f"Could not find or position window: {title}")

        # Show completion notification and exit
        self.toaster.show_toast(
            self.TOAST_TITLE,
            f"Completed! Successfully organized {success_count} of {total_windows} windows",
            duration=5,
            threaded=False
        )
        
        # Log completion and exit
        self.logger.info("Organization process completed. Exiting...")

def main():
    organiser = WindowOrganiser()
    organiser.organize_windows()

if __name__ == "__main__":
    main()
