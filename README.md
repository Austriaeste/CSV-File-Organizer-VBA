# CSV File Organizer VBA

## Overview
This VBA script automates the organization of CSV files based on their embedded date information. It replicates batch file functionality in VBA, sorting CSV files from a source folder into a hierarchical structure in a target folder (e.g., OneDrive\Desktop) based on year, 10-day periods, and 2-day intervals. The script includes robust error handling and a recursive folder creation function to prevent common issues like Error 76 (Path Not Found).

## Features
- **Date-Based Sorting**: Organizes CSV files (e.g., `2025050900_2025051023_XXX.csv`) into folders by year, 10-day periods (e.g., `2025-0501-0510`), and 2-day intervals (e.g., `0501_0502_収集ファイル`).
- **Safe Folder Creation**: Uses `CreateFolderRecursive` to create nested directories, avoiding Error 76.
- **Temporary Processing**: Processes files in a temporary folder before copying to the final destination, ensuring clean operations.
- **Recursive Copying**: Replaces `robocopy` with a VBA-based recursive folder copy function.
- **Error Handling**: Includes comprehensive error handling with user-friendly messages.

## Prerequisites
- **Microsoft Excel**: The script is designed to run in Excel's VBA environment.
- **FileSystemObject**: Uses `Scripting.FileSystemObject` for file and folder operations.
- **OneDrive**: Assumes files are stored in a OneDrive folder (adjustable in the script).
- **Windows OS**: Relies on Windows environment variables (e.g., `USERPROFILE`).

## Folder Structure
The script expects and creates the following folder structure:
- **Source Folder**: `%USERPROFILE%\Downloads\収集ファイル` (contains input CSV files).
- **Temporary Folder**: `%USERPROFILE%\Downloads\収集一時格納` (temporary processing folder).
- **Target Folder**: `%USERPROFILE%\OneDrive\デスクトップ\収集データ` (final storage location).

The final structure under the target folder looks like:
```
収集データ
└── 2025年度
    └── 2025-0501-0510
        └── 0501_0502_収集ファイル
            └── 2025050900_2025051023_XXX.csv
```

## Installation
1. **Open Excel**: Create or open an Excel workbook.
2. **Access VBA Editor**: Press `Alt + F11` to open the VBA editor.
3. **Insert Module**: Insert a new module and paste the script code.
4. **Adjust Paths** (if needed): Modify the `oneDrivePath` variable in the `OrganizeCollectedData` subroutine if your OneDrive path differs from the default (`%USERPROFILE%\OneDrive`).
5. **Save Workbook**: Save the workbook with macro support (`.xlsm` format).

## Usage
1. **Prepare CSV Files**: Place CSV files in the source folder (`%USERPROFILE%\Downloads\収集ファイル`). Files should follow the naming convention `YYYYMMDDHH_*.csv` (e.g., `2025050900_2025051023_XXX.csv`).
2. **Run the Macro**:
   - Press `Alt + F8`, select `OrganizeCollectedData`, and click **Run**.
   - Alternatively, assign the macro to a button in Excel.
3. **Check Output**: The script will:
   - Validate the source folder.
   - Create a temporary folder structure.
   - Sort and copy CSV files to the temporary folder.
   - Copy the organized structure to the final target folder.
   - Delete the temporary folder.
   - Display progress and error messages via message boxes.

## Script Details
- **Main Procedure**: `OrganizeCollectedData` orchestrates the entire process.
- **Helper Functions**:
  - `DeterminePeriodFolder`: Calculates 10-day period folders (e.g., `2025-0501-0510`).
  - `DetermineTwoDayFolder`: Determines 2-day interval folders (e.g., `0501_0502_収集ファイル`).
  - `RecursiveCopyFolder`: Recursively copies files and folders (replaces `robocopy`).
  - `CreateFolderRecursive`: Safely creates nested folder structures.
- **Error Handling**: Catches and reports errors, ensuring graceful failure with cleanup.

## Notes
- **File Naming**: The script expects CSV files with a date prefix in `YYYYMMDDHH` format (10 digits). Invalid filenames are skipped with debug output.
- **OneDrive Path**: If your OneDrive path differs, update the `oneDrivePath` variable in the script.
- **Debugging**: Check the VBA Immediate Window (`Ctrl + G`) for detailed logs of processed files and errors.
- **Temporary Folder**: The script deletes the temporary folder after processing. Ensure no other processes are using it.

## Limitations
- Designed for Windows environments with OneDrive.
- Assumes CSV files follow the specified naming convention.
- Requires manual adjustment for non-standard OneDrive paths.

## Contributing
Feel free to fork this repository, submit issues, or create pull requests for improvements. Suggestions for optimizing folder creation, error handling, or extending functionality are welcome!

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
