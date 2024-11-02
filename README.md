# Smartsheet Offload Utility

🚀 **Smartsheet Offload Utility** is a powerful tool designed to simplify the process of downloading and managing data from Smartsheet. With a user-friendly GUI and real-time progress updates, this utility is perfect for streamlining your workflow and keeping your data organized. 

## Features

- **Easy Smartsheet Data Downloads**: Download data from multiple Smartsheet sheets with a single click.
- **GUI Interface**: User-friendly interface built with `tkinter`.
- **Real-Time Progress**: See download progress in real-time with a visual progress bar.
- **Log Output**: View detailed log information, color-coded for easy readability.
- **Customizable**: Update access tokens, sheet IDs, and file paths easily.

## Installation

1. **Clone the Repository**:
    ```sh
    git clone https://github.com/yourusername/SmartsheetDownloader.git
    cd SmartsheetDownloader
    ```

2. **Install Required Packages**:
    ```sh
    pip install -r requirements.txt
    ```

## Usage

1. **Run the Application**:
    ```sh
    python your_script.py
    ```

2. **Fill in the Required Inputs**:
    - Smartsheet API access token
    - Smartsheet sheet IDs (comma-separated)
    - Directory path to save the Excel files

## Building an Executable

To build a standalone executable using PyInstaller:
```sh
pyinstaller --onefile --windowed your_script.py
