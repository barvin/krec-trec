# Excel File Processor - User Guide

## Quick Start

### Running the Application

1. **Locate the file**: `dist\krec_trec.exe`
2. **Double-click** to run the application
3. A window will open with a drag-and-drop area

### Processing a File

**Method 1: Drag and Drop**
1. Drag your Excel file from Windows Explorer
2. Drop it onto the application window
3. Wait for processing to complete
4. Click "Yes" to open the folder with the processed file

**Method 2: Browse**
1. Click anywhere in the white drop area
2. Select your Excel file from the file browser
3. Wait for processing to complete
4. Click "Yes" to open the folder with the processed file

### Output

- The processed file will be saved in the **same folder** as the input file
- The output filename will have **"_processed"** added before the extension
- Example: `21_05_2025_1690_SMA_SCID.xlsx` → `21_05_2025_1690_SMA_SCID_processed.xlsx`

## What Gets Processed

The application automatically:

✅ Finds the data table starting with "Well Position" header  
✅ Removes all metadata rows above the table  
✅ Removes unnecessary columns: Target Color, CQCONF, EXPFAIL, NOAMP  
✅ Adds calculated columns: RNP/2, Copies/mln, Delta  

### Calculations

For each sample:
- **RNP/2**: Takes the RNP Quantity value and divides it by 2
- **Copies/mln (KREC)**: RNP/2 minus KREC Quantity
- **Copies/mln (TREC)**: RNP/2 minus TREC Quantity

## Deploying to Another Computer

1. **Copy** `krec_trec.exe` to the target computer
2. **That's it!** No installation needed
3. Works on any Windows 10/11 computer

### System Requirements

- Windows 10 or Windows 11
- No Python installation required
- No additional software needed

## Troubleshooting

### Application won't start
- Ensure you're running on Windows 10 or 11
- Try running as Administrator (right-click → Run as Administrator)

### Processing fails
- Check that the Excel file contains a "Well Position" header
- Ensure the file is not corrupted or password-protected
- File format should be .xlsx or .xls

### File already exists error
- The output file already exists in the folder
- Delete or rename the previous processed file
- Or move your input file to a different folder

## Tips

💡 **Batch Processing**: Process multiple files one after another by dropping them sequentially

💡 **Keep Original**: The original file is never modified - a new file is always created

💡 **Fast Processing**: Most files process in just a few seconds

## Support

For issues or questions, contact your IT administrator or the application developer.

---

**Version**: 1.0.0  
**Last Updated**: January 2026
