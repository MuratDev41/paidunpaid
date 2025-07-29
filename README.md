# Paid/Unpaid Excel Classifier

This tool processes an Excel file to classify rows as "Paid" or "Unpaid" based on matching positive and negative price values, and adds a running count column. It features a graphical interface for selecting the input file and choosing which columns to use for price, count, and paid/unpaid status.

## Features
- Select any Excel file to process
- Choose which column contains the price data
- Choose which column to use for the running count (will be created or overwritten)
- Choose which column to use for the Paid/Unpaid result (will be created or overwritten)
- Fast processing with progress bar
- Results saved as a new Excel file in the same directory

## How It Works
For each row, the tool:
- Calculates a running count of each price value up to that row
- For any nonzero price, checks if the opposite value (e.g., -100 for 100) with the same count exists in the file
- Marks as "Paid" if a match is found, otherwise "Unpaid"

## Requirements
- Python 3.7+
- `pandas`, `tqdm`, `openpyxl` (for Excel support), `tkinter` (comes with most Python installations)

Install dependencies with:
```sh
pip install pandas tqdm openpyxl
```

## Usage
1. Run the script:
   ```sh
   python calculate_paid_unpaid.py
   ```
2. Use the file dialog to select your Excel file.
3. Enter/select the column names for price, count, and paid/unpaid when prompted.
4. Wait for processing to complete. The result will be saved as `<original_filename>_with_paid_unpaid.xlsx` in the same folder.
5. A pop-up will confirm when the process is done.

## Building a Standalone Executable
You can create a Windows executable using PyInstaller:

1. Install PyInstaller:
   ```sh
   pip install pyinstaller
   ```
2. Build the executable:
   ```sh
   pyinstaller --onefile --noconsole calculate_paid_unpaid.py
   ```
3. The executable will be in the `dist` folder as `calculate_paid_unpaid.exe`.

## Notes
- The script works with `.xlsx` and `.xls` files.
- The count and paid/unpaid columns will be created or overwritten as specified.
- If you select a column name that already exists, its contents will be replaced.
- The GUI uses pop-up dialogs for all user input and notifications.

## License
MIT License 