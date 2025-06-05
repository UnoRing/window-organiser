import os
import sys
import traceback
import time
import ctypes
from ctypes import wintypes
import subprocess
import shlex

try:
    import pygetwindow as gw
    import json
    import AppOpener
    import logging
    from pathlib import Path
    from win10toast import ToastNotifier

    # Setup logging first thing so we can log any errors
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
    logger = logging.getLogger("OrganizerProcess")
    logger.info("=== New Organization Process Started ===")

    def force_exit():
        """Force terminate the current process using Windows API"""
        try:
            pid = os.getpid()
            handle = ctypes.windll.kernel32.OpenProcess(1, False, pid)
            ctypes.windll.kernel32.TerminateProcess(handle, 0)
            ctypes.windll.kernel32.CloseHandle(handle)
        except Exception:
            # If Windows API fails, fallback to os._exit
            os._exit(0)

    class WindowOrganiser:
        TOAST_TITLE = "Window Organizer"
        MAX_UI_WAIT_TIME = 30  # Maximum seconds to wait for UI

        def __init__(self):
            self.logger = logging.getLogger("OrganizerProcess")  # Use the global logger
            self.config_file = "window_config.json"
            self.load_config()
            
            # Wait for UI to be ready before initializing toaster
            if not self.wait_for_ui_ready():
                self.logger.error("UI not ready after timeout, exiting...")
                force_exit()
            
            self.toaster = ToastNotifier()

            self.logger.info("Starting organization process...")

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

        def launch_and_wait_for_window(self, title, app_name, original_title=None, max_attempts=10, delay=1, opening_method=None):
            """Try to launch an app and wait for its window to appear."""
            if opening_method:
                self.logger.info(f"Launching {title} using custom opening_method: {opening_method}")
                try:
                    cmd_list = shlex.split(opening_method, posix=False)
                    subprocess.Popen(cmd_list)
                except Exception as e:
                    self.logger.error(f"Failed to launch {title} with subprocess: {e}")
                    return None
                title_variations = [title]
                if original_title:
                    title_variations.append(original_title)
                max_wait = max_attempts
            else:
                launch_methods = [
                    lambda: AppOpener.open(app_name, match_closest=True),
                    lambda: AppOpener.open(app_name.replace(" ", "-"), match_closest=True),
                    lambda: AppOpener.open(app_name.replace(" ", ""), match_closest=True),
                    lambda: os.system(f"start {app_name}")
                ]
                title_variations = [title]
                if original_title:
                    title_variations.append(original_title)
                self.logger.info(f"Attempting to launch {app_name} (original title: {original_title})")
                launch_success = False
                for i, method in enumerate(launch_methods):
                    try:
                        self.logger.info(f"Launch attempt {i+1} with method: {method.__name__ if hasattr(method, '__name__') else 'lambda'}")
                        method()
                        all_windows = gw.getAllWindows()
                        window_titles = [w.title for w in all_windows if w.title]
                        self.logger.info(f"Current windows after launch attempt: {window_titles}")
                        launch_success = True
                        break
                    except Exception as e:
                        self.logger.warning(f"Launch attempt {i+1} failed: {e}")
                        continue
                if not launch_success:
                    self.logger.error(f"All launch attempts failed for {app_name}")
                    return None
                max_wait = max_attempts

            # Wait for window to appear
            for attempt in range(max_wait):
                time.sleep(delay)
                for title_var in title_variations:
                    try:
                        windows = gw.getWindowsWithTitle(title_var)
                        if windows:
                            self.logger.info(f"Window appeared after {attempt + 1}s with title: {title_var}")
                            return windows[0]
                    except Exception as e:
                        self.logger.warning(f"Error getting window {title_var}: {e}")
                if attempt == 0:
                    all_windows = gw.getAllWindows()
                    window_titles = [w.title for w in all_windows if w.title]
                    self.logger.info(f"Available windows while waiting: {window_titles}")
                self.logger.info(f"Waiting... ({attempt + 1}/{max_wait})")
            self.logger.warning(f"Window did not appear after {max_wait} attempts")
            return None

        def position_window(self, window, config, title):
            """Position window with a single attempt and simple verification"""
            try:
                # Get initial position
                initial_pos = {
                    'left': window.left,
                    'top': window.top,
                    'width': window.width,
                    'height': window.height
                }
                
                # Try to position
                try:
                    max_attempts = 1
                    for attempt in range(max_attempts):
                        try:
                            if attempt > 0:
                                # Get fresh handle for retry
                                windows = gw.getWindowsWithTitle(title)
                                if not windows:
                                    self.logger.warning(f"Window disappeared during retry {attempt}")
                                    return False
                                window = windows[0]
                                self.logger.info(f"Retry {attempt} with fresh handle: {window._hWnd}")
                                time.sleep(0.5)  # Small delay between retries
                                
                            window.resizeTo(config["width"], config["height"])
                            window.moveTo(config["x"], config["y"])
                            self.logger.info(f"Positioned: {title}")
                            break
                        except Exception as e:
                            if attempt == max_attempts - 1:  # Last attempt
                                raise  # Re-raise the exception if all attempts failed
                            self.logger.warning(f"Position attempt {attempt + 1} failed: {e}")
                            continue
                except Exception as e:
                    self.logger.error(f"Failed to position {title}: {e}")
                    self.logger.error(f"Window state at failure - x={initial_pos['left']}, y={initial_pos['top']}, w={initial_pos['width']}, h={initial_pos['height']}")
                    return False
                
                # Small delay to let window settle
                time.sleep(0.2)
                
                # Simple verification with tolerance
                try:
                    # Get fresh window handle
                    windows = gw.getWindowsWithTitle(title)
                    if not windows:
                        self.logger.warning(f"Window disappeared after positioning: {title}")
                        return False
                        
                    window = windows[0]
                    current_pos = {
                        'left': window.left,
                        'top': window.top,
                        'width': window.width,
                        'height': window.height
                    }
                    
                    # Check each dimension separately
                    pos_ok = True
                    mismatches = []
                    
                    if abs(current_pos['left'] - config["x"]) > 2:
                        mismatches.append(f"X: expected {config['x']}, got {current_pos['left']}")
                        pos_ok = False
                    if abs(current_pos['top'] - config["y"]) > 2:
                        mismatches.append(f"Y: expected {config['y']}, got {current_pos['top']}")
                        pos_ok = False
                    if abs(current_pos['width'] - config["width"]) > 2:
                        mismatches.append(f"Width: expected {config['width']}, got {current_pos['width']}")
                        pos_ok = False
                    if abs(current_pos['height'] - config["height"]) > 2:
                        mismatches.append(f"Height: expected {config['height']}, got {current_pos['height']}")
                        pos_ok = False
                    
                    if pos_ok:
                        return True
                    else:
                        self.logger.warning(f"Position verification failed for {title}:")
                        self.logger.warning(f"Initial position: x={initial_pos['left']}, y={initial_pos['top']}, w={initial_pos['width']}, h={initial_pos['height']}")
                        self.logger.warning(f"Target position: x={config['x']}, y={config['y']}, w={config['width']}, h={config['height']}")
                        self.logger.warning(f"Final position: x={current_pos['left']}, y={current_pos['top']}, w={current_pos['width']}, h={current_pos['height']}")
                        self.logger.warning(f"Mismatches: {', '.join(mismatches)}")
                        return False
                        
                except Exception as e:
                    self.logger.warning(f"Verification failed for {title}: {e}")
                    # If verification fails but initial position was different, assume success
                    if (initial_pos['left'] != config["x"] or 
                        initial_pos['top'] != config["y"] or 
                        initial_pos['width'] != config["width"] or 
                        initial_pos['height'] != config["height"]):
                        self.logger.info(f"Assuming success for {title} based on position change from initial state")
                        return True
                    return False
                    
            except Exception as e:
                self.logger.error(f"Failed to position {title}: {e}")
                return False

        def organize_windows(self):
            """Open and organize all windows according to saved configuration"""
            try:
                self.toaster.show_toast(
                    self.TOAST_TITLE,
                    "Starting to organize windows...",
                    duration=3,
                    threaded=False
                )
                self.logger.info("Starting window organization")
                success_count = 0
                total_windows = len(self.window_configs)
                failed_apps = []  # Track failed apps

                for title, config in self.window_configs.items():
                    try:
                        self.logger.info(f"Processing: {title}")
                        windows = gw.getWindowsWithTitle(title)
                        position_only = config.get("position_only", False)
                        open_method = config.get("open_method", "")
                        opening_method = config.get("opening_method", "")
                        if position_only:
                            self.logger.info(f"position_only set for {title}: will not launch, just wait for window (up to 30s)")
                            for attempt in range(30):
                                if windows:
                                    break
                                time.sleep(1)
                                windows = gw.getWindowsWithTitle(title)
                                self.logger.info(f"Waiting for window... ({attempt+1}/30)")
                        else:
                            # Use opening_method if specified
                            if not windows:
                                window = None
                                if opening_method:
                                    window = self.launch_and_wait_for_window(
                                        title, config["app_name"], config.get("original_title"),
                                        opening_method=opening_method
                                    )
                                elif open_method == "appopener":
                                    try:
                                        AppOpener.open(config["app_name"], match_closest=True)
                                    except Exception as e:
                                        self.logger.warning(f"AppOpener launch failed: {e}")
                                elif open_method == "system":
                                    os.system(f"start {config['app_name']}")
                                elif open_method.startswith("custom:"):
                                    cmd = open_method[len("custom:"):]
                                    os.system(cmd)
                                elif open_method:
                                    self.logger.warning(f"Unknown open_method '{open_method}' for {title}, falling back to default.")
                                else:
                                    # Default: use current launch logic
                                    window = self.launch_and_wait_for_window(
                                        title, config["app_name"], config.get("original_title")
                                    )
                                # Wait for window to appear
                                for attempt in range(10):
                                    time.sleep(1)
                                    windows = gw.getWindowsWithTitle(title)
                                    if not windows and "original_title" in config:
                                        windows = gw.getWindowsWithTitle(config["original_title"])
                                    if windows:
                                        self.logger.info(f"Window appeared after {attempt + 1}s")
                                        break
                                    self.logger.info(f"Waiting... ({attempt + 1}/10)")
                            else:
                                window = windows[0]
                        if position_only:
                            window = windows[0] if windows else None
                            if window:
                                if self.position_window(window, config, title):
                                    success_count += 1
                                else:
                                    failed_apps.append(title)
                            else:
                                self.logger.warning(f"Could not find window: {title}")
                                failed_apps.append(title)
                        else:
                            if window:
                                if self.position_window(window, config, title):
                                    success_count += 1
                                else:
                                    failed_apps.append(title)
                            else:
                                self.logger.warning(f"Could not find or position window: {title}")
                                failed_apps.append(title)
                    except Exception as e:
                        self.logger.error(f"Error processing window {title}: {e}")
                        failed_apps.append(title)
                        continue

                # Prepare toast message
                if not failed_apps:
                    toast_message = "Window-Organiser completed successfully"
                else:
                    toast_message = f"Window-Organiser completed successfully\nMinor issue with: {', '.join(failed_apps)}"

                self.toaster.show_toast(
                    self.TOAST_TITLE,
                    toast_message,
                    duration=5,
                    threaded=False
                )
                self.logger.info(f"Organization process completed. Success: {success_count}/{total_windows}")
                if failed_apps:
                    self.logger.info(f"Failed apps: {', '.join(failed_apps)}")
            except Exception as e:
                self.logger.error(f"Critical error during organization: {e}")
                traceback.print_exc(file=sys.stderr)
            finally:
                force_exit()

        def wait_for_ui_ready(self):
            """Wait for UI to be ready, with timeout"""
            start_time = time.time()
            while time.time() - start_time < self.MAX_UI_WAIT_TIME:
                try:
                    # Try to get the desktop window as a test
                    desktop = ctypes.windll.user32.GetDesktopWindow()
                    if desktop:
                        self.logger.info("UI is ready")
                        return True
                except Exception:
                    pass
                time.sleep(1)
                self.logger.info("Waiting for UI to be ready...")
            
            self.logger.warning("UI ready timeout reached after 30 seconds")
            return False

    def main():
        try:
            organiser = WindowOrganiser()
            organiser.organize_windows()
        except Exception:
            logger.error("Critical error in main:")
            logger.error(traceback.format_exc())
            force_exit()

    if __name__ == "__main__":
        main()
except Exception as e:
    # If we can't even setup logging, write to a file directly
    with open('organizer_process.log', 'a', encoding='utf-8') as f:
        f.write("\n=== Critical Startup Error ===\n")
        f.write(traceback.format_exc())
    force_exit()
