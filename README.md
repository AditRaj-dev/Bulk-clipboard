# Excel Clipboard Automation

A Windows desktop app for automating Excel clipboard tasks using a GUI built with Tkinter. Includes features for managing Excel files, clipboard operations, and automation scripts.

## Features
- Excel file management (openpyxl)
- Clipboard automation (pyperclip, keyboard)
- GUI interface (Tkinter)
- Automation scripts and helpers

## Installation
1. Clone or download this repository.
2. Install dependencies:
   - Recommended: Run the install script:
     ```sh
     python install_deps.py
     ```
   - Or manually:
     ```sh
     pip install -r requirements.txt
     ```

## Usage
Start the application:
```sh
python app.py
```

## Packaging as EXE
To create a standalone Windows executable:
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
