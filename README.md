# 📋 Excel Clipboard Automation

A **Windows desktop application** for automating clipboard-based workflows using **Excel**, built with **Python and Tkinter**, with optional **AutoHotkey (AHK)** integration for full automation.

This tool is designed to reduce repetitive manual work such as copying, pasting, submitting, and navigating through large sets of questions or code snippets.

---

## 🚀 Features

### 🔹 Excel Management

* Uses a single Excel file (`.xlsx`) as persistent storage
* Supports **multiple sheets** for different subjects or exams
* Automatically creates the Excel file and sheets if missing
* Tracks cursor position **per sheet**

---

### 🔹 Clipboard Capture (Clipboard → Excel)

* Start capturing clipboard contents with one click
* Each new copy (`Ctrl + C`) is saved to the next Excel row
* Prevents duplicate consecutive captures
* Stop capture at any time

---

### 🔹 Clipboard Loading (Excel → Clipboard)

* Continuous loading mode
* Load snippet → paste → auto-load next snippet
* Cursor advances automatically
* Safe stop and restart without losing position

---

### 🔹 Question-Based Loading (Advanced Mode)

Designed for workflows where question numbers don’t start at 1 (e.g. 1676, 1677, …).

User can enter:

* **Start Question**
* **End Question**
* **From Question**
* **Base Question** (question number corresponding to Excel row 1)

The app automatically:

* Calculates the correct Excel start row
* Calculates how many questions remain
* Calculates required AHK loop count

---

### 🔹 AutoHotkey Automation (Optional)

* Integrates with **AutoHotkey v1.1**
* Supports:

  * Capture automation
  * Loading / submit / next-question automation
* Hotkey-based triggers
* Emergency stop using `Esc`

---

### 🔹 Clean GUI (Dark Mode)

* Built with Tkinter
* Grouped sections for clarity
* Advanced controls hidden by default
* Real-time activity logs

---

## 🛠 Installation

### 1️⃣ Prerequisites

* **Windows OS**
* **Python 3.8+**
* **AutoHotkey v1.1** (optional, for automation)

Check Python:

```sh
python --version
```

---

### 2️⃣ Install Python Dependencies

```sh
pip install openpyxl pyperclip keyboard
```
**OR**
run the ```sh
install_deps.py
``` file

> ⚠️ The `keyboard` library may require administrator privileges on Windows.

---

### 3️⃣ Install AutoHotkey (Optional)

Download and install **AutoHotkey v1.1** from:
👉 [https://www.autohotkey.com/](https://www.autohotkey.com/)

> This project uses **AHK v1 syntax**, not v2.

---

## ▶ Running the Application

Start the application using:

```sh
python app.py
```

The main GUI window will open automatically.

---

## 🧭 How to Use the App

### 🟢 1. Select Context

* Launch the app
* Select the Excel sheet you want to work with
* View the current cursor row (read-only)

---

### 🟢 2. Capture Mode (Clipboard → Excel)

1. Click **Start Capture**
2. Copy content normally (`Ctrl + C`)
3. Each new clipboard entry is saved to Excel
4. Click **Stop Capture** when finished

Useful for collecting answers or code snippets.

---

### 🟢 3. Loading Mode (Excel → Clipboard)

#### Simple Mode (Row-Based)

1. Enter the **Start Row**
2. Click **Start Loading**
3. Paste using `Ctrl + V`
4. Cursor advances automatically
5. Next snippet loads automatically

---

#### Advanced Mode (Question-Based)

1. Enable **Advanced Question Mode**
2. Enter:

   * Start Question
   * End Question
   * From Question
   * Base Question
3. The app calculates:

   * Excel start row
   * Total questions
   * AHK loop count
4. Click **Start Loading**

Ideal when question numbering is non-linear.

---

### 🟢 4. AutoHotkey Automation

* Trigger AHK actions from the UI
* Or run AHK scripts directly
* Press **Esc** anytime to emergency-stop automation

---

### 🟢 5. Logs

* All actions are logged in real time
* Use **Clear Logs** to reset the log panel

---

## ⌨ Default Hotkeys

| Action          | Hotkey       |
| --------------- | ------------ |
| Paste Detection | `Ctrl + V`   |
| AHK Triggers    | Configurable |
| Emergency Stop  | `Esc`        |

Hotkeys can be modified in the AHK templates or `config.py`.

---

## 📦 Packaging as EXE (Windows)

To create a standalone executable:

### 1️⃣ Install PyInstaller

```sh
pip install pyinstaller
```

### 2️⃣ Build the EXE

```sh
pyinstaller --onefile app.py
```

The executable will be available in the `dist/` folder.

---

## 📁 Project Structure

```
project/
│
├── app.py                # Main entry point
├── ui.py                 # GUI layout
├── automation.py         # Capture & loading logic
├── excel_manager.py      # Excel file operations
├── clipboard_manager.py  # Clipboard helpers
├── ahk_manager.py        # AutoHotkey integration
├── logger.py             # Logging utility
├── state.py              # Shared state flags
├── config.py             # Constants & settings
│
├── template_capture.ahk
├── template_loading.ahk
│
└── java_snippets.xlsx    # Auto-created Excel file
```

---

## ⚠ Notes & Tips

* Cursor position is tracked **per Excel sheet**
* Avoid running multiple AHK instances simultaneously
* Always verify **Base Question** when using advanced mode
* Generated AHK scripts overwrite previous versions

---

## 📜 License

MIT License

---

## ✅ Summary

This application turns a repetitive, manual clipboard workflow into a **controlled, reliable automation system**, while still keeping the user in full control.

Whether you are using it for exams, practice platforms, or bulk data entry, this tool is built to **save time and reduce errors**.

---

If you want next:

* Screenshots section
* Quick-start TL;DR
* Troubleshooting guide
* Demo GIF walkthrough

Just say the word 👍
