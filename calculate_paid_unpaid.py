import pandas as pd
from tqdm import tqdm
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

# File selection
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename(
    title="Select Excel file",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)
if not file_path:
    print("No file selected. Exiting.")
    exit()
print("Please wait while the file is being loaded...")
df = pd.read_excel(file_path)

# Ask user which columns to use
columns = list(df.columns)

# Ask for price data column
price_col = simpledialog.askstring(
    "Select Price Column",
    f"Available columns: {columns}\nEnter the column name to use for price data:",
    initialvalue=columns[2] if len(columns) > 2 else columns[0]
)
if price_col not in columns:
    messagebox.showerror("Error", f"Column '{price_col}' not found.")
    exit()

# Ask for count column (where to add the running count)
count_col = simpledialog.askstring(
    "Select Count Column",
    f"Available columns: {columns}\nEnter the column name to use for the count (will be overwritten or created):",
    initialvalue="D"
)
if not count_col:
    messagebox.showerror("Error", "No count column specified.")
    exit()

# Ask for paid/unpaid column (where to add the result)
paid_col = simpledialog.askstring(
    "Select Paid/Unpaid Column",
    f"Available columns: {columns}\nEnter the column name to use for Paid/Unpaid (will be overwritten or created):",
    initialvalue="E"
)
if not paid_col:
    messagebox.showerror("Error", "No paid/unpaid column specified.")
    exit()

# Calculate count_col: running count of each value in price_col
df[count_col] = df.groupby(price_col).cumcount() + 1

# Build a set of (price, count) pairs for fast lookup
cd_set = set(zip(df[price_col], df[count_col]))

# Prepare paid_col
E = []
for idx, row in tqdm(df.iterrows(), total=len(df), desc="Classifying Paid/Unpaid"):
    c_val = row[price_col]
    d_val = row[count_col]
    if c_val != 0:
        E.append("Paid" if (-c_val, d_val) in cd_set else "Unpaid")
    else:
        E.append("")
df[paid_col] = E

output_file = file_path.rsplit('.', 1)[0] + '_with_paid_unpaid.xlsx'
print("Please wait while the file is being saved...")
df.to_excel(output_file, index=False)
print(f"Saved results to {output_file}")
messagebox.showinfo("Done", f"Saved results to {output_file}")
