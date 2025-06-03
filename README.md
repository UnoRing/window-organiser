# Window Organizer

A Python automation tool to open and organize your favorite Windows apps and windows with a single click or at startup.

## Features
- Automatically opens and arranges your selected apps/windows
- Saves and restores window positions and sizes
- Modern GUI for configuration
- Toast notifications with custom icon support
- Logging for troubleshooting

## Requirements
- Windows 10/11
- Python 3.8+
- The following Python packages:
  - pygetwindow
  - AppOpener
  - win10toast
  - logging (standard library)

Install requirements:
```bash
pip install pygetwindow AppOpener win10toast
```

## Setup
1. **Clone or download this repository.**
2. **Configure your windows:**
   - Run `window-organiser-config.py` to open the GUI.
   - Arrange your windows as desired, select them, and save their positions.
   - The configuration is saved in `window_config.json`.
3. **(Optional) Add a custom icon:**
   - Place a `.ico` file (e.g., `window-organiser-icon.ico`) in the project directory.
   - The script will use this icon for toast notifications.

## Usage
- To organize your windows, run:
  ```bash
  python window-organiser.py
  ```
- You will see a toast notification that the script is waiting 20 seconds (to allow system startup).
- After 20 seconds, the script will open and arrange your windows, showing progress via toasts.
- All actions are logged in `organizer_process.log`.

## Troubleshooting
- **Toast icon not showing?**
  - Ensure your icon is a valid `.ico` file (not `.png` or `.jpg`).
  - Use a relative path like `window-organiser-icon.ico` if the icon is in the same directory.
  - Icon should be 32x32 or 64x64 pixels for best results.
- **Errors about `stderr` or console:**
  - Run the script as `.py` (not `.pyw`) for debugging.
  - All output is logged in `organizer_process.log`.
- **Some windows/apps not opening or positioning?**
  - Make sure the app name mapping in the config matches what AppOpener expects.
  - Update all dependencies to the latest version.

## License
MIT 