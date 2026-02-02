# Excel File Processor

A Windows desktop application for processing Excel files from laboratory analysis. The application provides a drag-and-drop interface for easy file processing.

## Features

- **Drag-and-Drop Interface**: Simple and intuitive UI
- **Automatic Processing**: Finds and processes data tables automatically
- **Column Management**: Removes unnecessary columns and adds calculated fields
- **Standalone Executable**: Single .exe file that works on any Windows computer

## What the Application Does

The application processes Excel files by:

1. Finding the data table that starts with "Well Position" header
2. Removing all rows above the data table
3. Removing columns: "Target Color", "CQCONF", "EXPFAIL", "NOAMP"
4. Adding new columns: "RNP/2", "Copies/mln", "Delta"
5. Calculating values:
   - **RNP/2**: For each sample, divides the RNP Quantity value by 2
   - **Copies/mln**: For KREC and TREC rows, calculates RNP/2 minus Quantity

## Installation & Setup

### Option 1: Use the Pre-built Executable (Recommended)

1. Copy `krec_trec.exe` to your computer
2. Double-click to run
3. No installation required!

### Option 2: Build from Source

1. Install Python 3.8 or higher
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Build the executable:
   ```bash
   python build_exe.py
   ```
4. The executable will be created in the `dist` folder

## Usage

1. **Run the Application**: Double-click `krec_trec.exe`
2. **Process a File**: 
   - Drag and drop an Excel file onto the window, OR
   - Click on the drop area to browse for a file
3. **Wait for Completion**: The app will process the file and show a success message
4. **Get Results**: The processed file will be saved in the same folder with "_processed" suffix

## File Structure

```
krec-trec/
├── excel_processor.py    # Core processing logic
├── gui_app.py           # GUI application
├── build_exe.py         # Build script for creating .exe
├── requirements.txt     # Python dependencies
└── README.md           # This file
```

## Requirements

- Windows operating system
- For running the .exe: No requirements (standalone)
- For development:
  - Python 3.8+
  - pandas
  - openpyxl
  - tkinterdnd2
  - pyinstaller (for building)

## Example

**Input File**: `21_05_2025_1690_SMA_SCID.xlsx`

**Output File**: `21_05_2025_1690_SMA_SCID_processed.xlsx`

The processed file will contain:
- Clean data table starting from the "Well Position" header
- Removed unnecessary columns
- Added calculated columns (RNP/2, Copies/mln, Delta)

## Development

To run the application without building:

```bash
python gui_app.py
```

To test the processing logic:

```python
from excel_processor import process_excel_file
output = process_excel_file("input_file.xlsx")
print(f"Processed file saved to: {output}")
```

## Troubleshooting

**Problem**: Application doesn't start
- Solution: Make sure you're running on Windows
- Try running in compatibility mode

**Problem**: File processing fails
- Solution: Ensure the Excel file contains a "Well Position" header
- Check that the file format is .xlsx or .xls

**Problem**: Missing dependencies during development
- Solution: Run `pip install -r requirements.txt`

## License

This project is for internal laboratory use.

## Version

1.0.0 - Initial Release
