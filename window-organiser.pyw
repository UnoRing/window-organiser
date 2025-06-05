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

        def wait_for_ui_ready(self):
            """Wait for UI to be ready, with timeout"""
            start_time = time.time()
            while time.time() - start_time < self.MAX_UI_WAIT_TIME:
                try:
                    # Try to get the desktop window as a test
                    desktop = ctypes.windll.user32.GetDesktopWindow()
                    if desktop:
                        logger.info("UI is ready")
                        return True
                except Exception:
                    pass
                time.sleep(1)
                logger.info("Waiting for UI to be ready...")
            
            logger.warning("UI ready timeout reached after 30 seconds")
            return False

        def __init__(self):
            self.setup_logging()
            self.config_file = "window_config.json"
            self.load_config()
            
            # Wait for UI to be ready before initializing toaster
            if not self.wait_for_ui_ready():
                logger.error("UI not ready after timeout, exiting...")
                force_exit()
            
            self.toaster = ToastNotifier()

            logger.info("Starting organization process...")

        def setup_logging(self):
            """Initialize logging"""
            logger.info("=== New Organization Process Started ===")

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
                logger.error(f"Error loading config: {e}")
                self.window_configs = {}

        def launch_and_wait_for_window(self, title, app_name, original_title=None, max_attempts=10, delay=1, opening_method=None):
            """Try to launch an app and wait for its window to appear."""
            if opening_method:
                logger.info(f"Launching {title} using custom opening_method: {opening_method}")
                try:
                    cmd_list = shlex.split(opening_method, posix=False)
                    subprocess.Popen(cmd_list)
                except Exception as e:
                    logger.error(f"Failed to launch {title} with subprocess: {e}")
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
                logger.info(f"Attempting to launch {app_name} (original title: {original_title})")
                launch_success = False
                for i, method in enumerate(launch_methods):
                    try:
                        logger.info(f"Launch attempt {i+1} with method: {method.__name__ if hasattr(method, '__name__') else 'lambda'}")
                        method()
                        all_windows = gw.getAllWindows()
                        window_titles = [w.title for w in all_windows if w.title]
                        logger.info(f"Current windows after launch attempt: {window_titles}")
                        launch_success = True
                        break
                    except Exception as e:
                        logger.warning(f"Launch attempt {i+1} failed: {e}")
                        continue
                if not launch_success:
                    logger.error(f"All launch attempts failed for {app_name}")
                    return None
                max_wait = max_attempts
            for attempt in range(max_wait):
                time.sleep(delay)
                for title_var in title_variations:
                    windows = gw.getWindowsWithTitle(title_var)
                    if windows:
                        logger.info(f"Window appeared after {attempt + 1}s with title: {title_var}")
                        return windows[0]
                if attempt == 0:
                    all_windows = gw.getAllWindows()
                    window_titles = [w.title for w in all_windows if w.title]
                    logger.info(f"Available windows while waiting: {window_titles}")
                logger.info(f"Waiting... ({attempt + 1}/{max_wait})")
            logger.warning(f"Window did not appear after {max_wait} attempts")
            return None

        def position_window(self, window, config, title):
            try:
                window.resizeTo(config["width"], config["height"])
                window.moveTo(config["x"], config["y"])
                logger.info(f"Positioned: {title}")
                return True
            except Exception as e:
                logger.error(f"Failed to position {title}: {e}")
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
                logger.info("Starting window organization")
                success_count = 0
                total_windows = len(self.window_configs)
                for title, config in self.window_configs.items():
                    try:
                        logger.info(f"Processing: {title}")
                        windows = gw.getWindowsWithTitle(title)
                        position_only = config.get("position_only", False)
                        open_method = config.get("open_method", "")
                        opening_method = config.get("opening_method", "")
                        if position_only:
                            logger.info(f"position_only set for {title}: will not launch, just wait for window (up to 30s)")
                            for attempt in range(30):
                                if windows:
                                    break
                                time.sleep(1)
                                windows = gw.getWindowsWithTitle(title)
                                logger.info(f"Waiting for window... ({attempt+1}/30)")
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
                                        logger.warning(f"AppOpener launch failed: {e}")
                                elif open_method == "system":
                                    os.system(f"start {config['app_name']}")
                                elif open_method.startswith("custom:"):
                                    cmd = open_method[len("custom:"):]
                                    os.system(cmd)
                                elif open_method:
                                    logger.warning(f"Unknown open_method '{open_method}' for {title}, falling back to default.")
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
                                        logger.info(f"Window appeared after {attempt + 1}s")
                                        break
                                    logger.info(f"Waiting... ({attempt + 1}/10)")
                            else:
                                window = windows[0]
                        if position_only:
                            window = windows[0] if windows else None
                            if window:
                                self.position_window(window, config, title)
                                # Check if position matches config AFTER positioning attempt
                                pos_ok = (
                                    window.left == config["x"] and
                                    window.top == config["y"] and
                                    window.width == config["width"] and
                                    window.height == config["height"]
                                )
                                if pos_ok:
                                    success_count += 1
                                    logger.info(f"Success for app: {title}")
                                else:
                                    logger.warning(f"Failure for app: {title} (actual: x={window.left}, y={window.top}, w={window.width}, h={window.height}; expected: x={config['x']}, y={config['y']}, w={config['width']}, h={config['height']})")
                            else:
                                logger.warning(f"Could not find window: {title}")
                                logger.warning(f"Failure for app: {title}")
                        else:
                            if window:
                                self.position_window(window, config, title)
                                # Check if position matches config AFTER positioning attempt
                                pos_ok = (
                                    window.left == config["x"] and
                                    window.top == config["y"] and
                                    window.width == config["width"] and
                                    window.height == config["height"]
                                )
                                if pos_ok:
                                    success_count += 1
                                    logger.info(f"Success for app: {title}")
                                else:
                                    logger.warning(f"Failure for app: {title} (actual: x={window.left}, y={window.top}, w={window.width}, h={window.height}; expected: x={config['x']}, y={config['y']}, w={config['width']}, h={config['height']})")
                            else:
                                logger.warning(f"Could not find or position window: {title}")
                                logger.warning(f"Failure for app: {title}")
                    except Exception as e:
                        logger.error(f"Error processing window {title}: {e}")
                        continue
                self.toaster.show_toast(
                    self.TOAST_TITLE,
                    f"Completed! Successfully organized {success_count} of {total_windows} windows",
                    duration=5,
                    threaded=False
                )
                logger.info(f"Organization process completed. Success: {success_count}/{total_windows}")
            except Exception as e:
                logger.error(f"Critical error during organization: {e}")
                traceback.print_exc(file=sys.stderr)
            finally:
                force_exit()

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
