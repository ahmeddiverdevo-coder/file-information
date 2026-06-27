# File Information Scanner

A simple Python desktop tool that scans a selected folder, extracts file metadata, and exports the results into a clean Excel report.

## Overview

File Information Scanner helps users create a quick inventory of files inside a selected folder and its subfolders.

The tool collects basic file information such as file name, full path, owner, creation time, and last modified time, then saves the results in an Excel file.

## Features

- Scan a selected folder and all its subfolders.
- Ask the user to select the folder path using a folder selection window.
- Extract file name and full file path.
- Get file owner information from Windows security metadata.
- Display file creation date and time.
- Display last modified date and time.
- Export results to an Excel `.xlsx` report.
- Auto-adjust Excel column widths for better readability.
- Simple graphical interface for folder selection and save location.

## Requirements

This script is designed to run on Windows because it uses Windows security metadata through `pywin32`.

Before running the script, install the required Python libraries:

```bash
pip install pandas pywin32 xlsxwriter
```

## How to Run the Script

1. Make sure Python is installed on your Windows machine.
2. Download or clone this repository.
3. Open Command Prompt or PowerShell inside the project folder.
4. Install the required libraries:

```bash
pip install pandas pywin32 xlsxwriter
```

5. Run the script:

```bash
python files_information.py
```

If your script has a different file name, replace `files_information.py` with the actual Python file name.

Example:

```bash
python "files information.py"
```

## How to Use

After running the script:

1. A folder selection window will open.
2. Select the folder path you want to scan.
3. The script will scan all files inside the selected folder and its subfolders.
4. After the scan is complete, a save window will appear.
5. Choose where to save the Excel report.
6. The results will be exported to an `.xlsx` file.

## Folder Path Selection

You do not need to type the folder path manually.

The script uses a graphical folder picker to ask the user to select the folder path.  
After selecting the folder, the script automatically scans that folder and all files inside its subfolders.

## Output Columns

The Excel report includes the following columns:

| Column | Description |
|---|---|
| File Name | The name of the file |
| Full Path | The complete location of the file |
| Owner | The Windows file owner |
| Created | File owner or fallback value with creation date and time |
| Last Modified | File owner or fallback value with last modified date and time |

## Important Note

This tool reads the current file metadata available from Windows.

The displayed owner usually represents the file owner, but it should not be considered a complete audit trail or legal proof of who created or modified the file.

To accurately track who modified a file, file auditing or logging should be enabled before the action occurs, such as:

- Windows File Auditing
- File Server Audit Logs
- SharePoint or OneDrive Version History
- Microsoft Purview Audit
- SIEM logging

## Use Cases

This tool can be useful for:

- File inventory reports
- Basic administrative file review
- Checking file ownership
- Reviewing creation and modification dates
- Preparing documentation for shared folders

## Limitations

- The tool may not accurately identify the real user who created or modified the file.
- The owner is not always the same as the original creator.
- It does not provide a full historical audit log.
- Accurate user activity tracking requires auditing or logging to be enabled before file changes occur.

## License

This project is open source. You may use and modify it according to your needs.
