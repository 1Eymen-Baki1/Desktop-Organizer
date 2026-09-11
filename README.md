# Desktop Organizer

A simple GUI application developed with Python that automatically organizes files on your desktop into folders based on their types.

## Features

- Automatic classification by file type (Images, Documents, Videos, Games, Applications)
- Custom category addition (user-defined extensions)
- Special detection for Steam and Riot Games:
  - Files with "steam" or "riot" in the name (case-insensitive)
  - .lnk shortcut files whose target contains "steam" or "riot" (e.g., Steam game shortcuts)
  - .url internet shortcut files whose content contains "steam" or "riot" (e.g., Steam:// URLs)
- Dark neon-themed modern interface
- One-click desktop organization
- Persistent custom categories via JSON configuration

## Requirements

- Python 3.6 or higher
- Tkinter (comes with Python)
- Optional: pywin32 (for .lnk target resolution, not required when running the bundled EXE; the EXE includes it)

## Usage

### As Python Script

1. Clone or download the repository.
2. Optionally install the optional dependency:
   ```bash
   pip install pywin32
   ```
3. Run `desktop_organizer.py`:
   ```bash
   python desktop_organizer.py
   ```
4. Click the "Organize Desktop" button in the application.

### As EXE (for Distribution) RECOMMEND

1. Download `dist\desktop_organizer.exe`.
2. Double-click to run (you may see a Windows SmartScreen warning; click "More info" → "Run anyway").
3. Click the "Organize Desktop" button.

## Building the EXE (for Developers)

Use PyInstaller to create a single-file EXE:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --hidden-import=win32com.client desktop_organizer.py
```

The EXE will be created in the `dist` folder as `desktop_organizer.exe`.

## Configuration

On first run, the application creates `organizer_config.json`. This file stores categories and their associated extensions. Custom categories added via the UI are saved to this file.

## License

This project is released under the MIT license for personal and educational use.

## Contact

For any issues, suggestions, or contributions, please open an issue.

---

Made with ❤️ by Claude Code

YOU CAN FİND .EXE FİLE İN THE DİST 
