import pandas as pd
import random

# Constants
NUM_ROWS = 1_000_000
OUTPUT_FILE = "random_data.xlsx"

# Price values you provided
price_values = [
    -30000.00,
    -33000.00,
    -1000.00,
    -39000.00,
    30000.00,
    -15000.00,
    1000.00
]

# Sample units
units = ['USD', 'EUR', 'TRY']

# Generate data
data = {
    "Document Number": [f"DOC{str(i).zfill(7)}" for i in range(1, NUM_ROWS + 1)],
    "Unit (Money)": [random.choice(units) for _ in range(NUM_ROWS)],
    "Price": [random.choice(price_values) for _ in range(NUM_ROWS)]
}

# Create DataFrame and write to Excel
df = pd.DataFrame(data)
df.to_excel(OUTPUT_FILE, index=False, engine='openpyxl')

print(f"Excel file with {NUM_ROWS} rows created as '{OUTPUT_FILE}'")
