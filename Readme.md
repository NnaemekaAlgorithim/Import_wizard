# Import Wizard

This is a GUI-based tool for importing data from Excel files into a SQL Server database.

## Table of Contents
- [Setting Up the Environment](#setting-up-the-environment)
- [Packaging for Windows](#packaging-for-windows)
- [Packaging for Linux](#packaging-for-linux)
- [Packaging for macOS](#packaging-for-macos)
- [Important Notes](#important-notes)

---

## Setting Up the Environment

To set up and run this application, follow these steps:

### Step 1: Create a Virtual Environment

Navigate to the project directory and create a virtual environment:
```bash
python3 -m venv venv
```
Creating a virtual environment is different for different operating systems, you can follow this youtube video for instructions on any of Linux, Windows or Mac [Video](https://youtu.be/kz4gbWNO1cw)

Activate the virtual environment:
* On Linux/macOS:

```bash
source venv/bin/activate
```

* On Windows:

```bash
venv\Scripts\activate
```

### Step 2: Install Dependencies
Navigate to the project directory and run the bash command:

```bash
pip install -r requirements.txt
```

---

## Packaging for Windows

### Step 1: Install PyInstaller

Ensure you have PyInstaller installed in your virtual environment:

```bash
pip install pyinstaller
```

### Step 2: Create Executable

Run the following command to package the application into a .exe file:

```bash
pyinstaller --onefile --noconsole import_wizard.py
```
* The --onefile option creates a single executable file.
* The --noconsole option suppresses the console window.

The packaged .exe file will be located in the dist folder. You can distribute this file as needed.

## Packaging for Linux

### Step 1: Install PyInstaller

Ensure you have PyInstaller installed in your virtual environment:

```bash
pip install pyinstaller
```

### Step 2: Create Executable

Run the following command to package the application:

```bash
pyinstaller --onefile import_wizard.py
```

The generated executable file will be located in the dist folder. You can run it using:

```bash
./dist/import_excel_to_sql
```

## Packaging for macOS

### Step 1: Install PyInstaller

Ensure you have PyInstaller installed in your virtual environment:

```bash
pip install pyinstaller
```

### Step 2: Create Executable

Run the following command to package the application:

```bash
pyinstaller --onefile import_wizard.py
```

The generated executable will be in the dist folder. Ensure it has the appropriate permissions to run:

```bash
chmod +x ./dist/import_wizard
```

If you encounter security warnings on macOS, sign the executable using:

```bash
codesign --force --deep --sign - ./dist/import_wizard
```

## Important Notes

* Ensure that the system where the application runs has ODBC Driver 17 for SQL Server installed and configured.

* The application relies on the presence of a valid Excel file with a consistent format to ensure smooth data import.

* If you encounter any dependency issues, verify your Python environment and reinstall packages using requirements.txt.