# Excel Clipboard Automation

A Windows desktop app for automating Excel clipboard tasks using a GUI built with Tkinter. Includes features for managing Excel files, clipboard operations, and automation scripts.




## Installation
```sh
python app.py
```

## Packaging as EXE
To create a standalone Windows executable:
## Usage
Start the application:
```sh
python app.py
```

## How to Operate
- Launch the app to open the main window.
- Use the interface to select or create Excel files.
- Use buttons and menus to copy, paste, and automate Excel clipboard tasks.
- Enter or paste data into fields as needed.
- Automation features allow quick row processing and clipboard management.
- Changes are saved automatically to the selected Excel file.
- Refer to on-screen instructions or tooltips for more help.
1. Install PyInstaller:
   ```sh
   pip install pyinstaller
   ```
2. Run:
   ```sh
   pyinstaller --onefile app.py
   ```
   The EXE will be in the `dist` folder.

## Project Structure
- app.py: Main entry point
- ui.py: GUI builder
- excel_manager.py: Excel file operations
- clipboard_manager.py: Clipboard helpers
- automation.py: Automation logic
- install_deps.py: Dependency installer
- requirements.txt: List of dependencies

## Requirements
- Python 3.8+
- Windows OS (recommended)

## License
MIT
