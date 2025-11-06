# Teams Assignment Downloader

A simple Python tool to download and organize Microsoft Teams assignment submissions from locally synced folders. Files are automatically renamed to each student's name while preserving file extensions.

## Features

- 🚀 **Simple & Interactive**: Prompts for all inputs (no command-line arguments required)
- 📁 **Local-Only**: Works directly with locally synced Teams/SharePoint folders - no API authentication needed
- 🎯 **Smart Assignment Detection**: Automatically detects assignment subfolders and lets you choose which to extract
- 📝 **Automatic Renaming**: Files are renamed to the student's display name while preserving extensions
- 🔄 **Duplicate Handling**: Automatically handles duplicate filenames with numbered suffixes

## Prerequisites

- Python 3.9 or higher
- Your Teams/SharePoint class site synced locally via OneDrive (see here for instructions: https://support.microsoft.com/en-gb/office/view-sharepoint-files-in-file-explorer-66b574bb-08b4-46b6-a6a0-435fd98194cc)
You can find your class's teams folder by going to Sharepoint>My Sites> click your team>Site Contents> Student Work>Submitted Files


## Installation

1. Clone this repository:
```bash
git clone https://github.com/06benste/teams-assignment-downloader.git
cd teams-assignment-downloader
```

2. Create a virtual environment (recommended):
```bash
python -m venv .venv
```

3. Activate the virtual environment:
   - **Windows (PowerShell):**
     ```bash
     .\.venv\Scripts\Activate.ps1
     ```
   - **Windows (CMD):**
     ```bash
     .venv\Scripts\activate.bat
     ```
   - **macOS/Linux:**
     ```bash
     source .venv/bin/activate
     ```

4. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

Run the script:
```bash
python TeamsAssignmentDownloader.py
```

The program will prompt you for:
1. **Submitted files folder path**: The root folder containing student submissions
   - Structure: `<Root>\<Student Name>\<Assignment Name>\<files>`
   - Example: `C:\Users\<you>\...\Class A (2024-2026) - Submitted files`
2. **Assignment selection**: Choose which assignment to extract from detected subfolders
3. **Output folder**: Where to save the renamed files (defaults to `downloads/<assignment_name>`)

## How It Works

1. Scans the submitted files folder for student directories
2. Detects assignment subfolders within each student's folder
3. Lets you select which assignment to extract
4. Copies all files from the selected assignment subfolders
5. Renames each file to the student's name (sanitized for Windows compatibility)
6. Saves files to the output folder, handling duplicates automatically

## Example

```
Input Structure:
  Submitted Files/
    ├── John Smith/
    │   └── Assignment 1/
    │       ├── report.docx
    │       └── data.xlsx
    └── Jane Doe/
        └── Assignment 1/
            └── report.docx

Output:
  downloads/Assignment 1/
    ├── John Smith.docx
    ├── John Smith.xlsx
    └── Jane Doe.docx
```

## Notes

- File names are sanitized to be Windows-compatible (invalid characters replaced with spaces, underscores converted to spaces)
- Duplicate filenames are automatically suffixed: `Name (2).ext`, `Name (3).ext`, etc.
- If a student doesn't have the selected assignment folder, they are skipped
- The output folder is automatically created if it doesn't exist

## Troubleshooting

- **Path not found**: Ensure the submitted files folder path is correct and accessible
- **No assignments detected**: Verify that student folders contain assignment subfolders
- **Permission errors**: Ensure you have read access to the source folder and write access to the output folder

## License

This project is open source and available under the [MIT License](LICENSE).

## Author

**06benste**

- GitHub: [@06benste](https://github.com/06benste)

## Contributing

Contributions, issues, and feature requests are welcome! Feel free to check the [issues page](https://github.com/06benste/teams-assignment-downloader/issues).

## Acknowledgments

Built for educators who need to quickly organize and review student assignment submissions from Microsoft Teams.

