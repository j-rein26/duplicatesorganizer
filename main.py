import pandas as pd

import re

# Read both files
df1 = pd.read_excel("NewHomeBuyersOct.xlsx")
df2 = pd.read_excel("NewMoversOct.xlsx")

# 2️⃣ Print column names to verify them
print("File 1 columns:", df1.columns.tolist())
print("File 2 columns:", df2.columns.tolist())

# Standardize columns
df1.columns = df1.columns.str.strip()
df2.columns = df2.columns.str.strip()

# Combine
combined = pd.concat([df1, df2], ignore_index=True)

# --- Clean and normalize addresses ---
def clean_address(addr):
    if pd.isna(addr):
        return ""
    addr = str(addr).strip().lower()
    addr = re.sub(r'\.', '', addr)        # remove periods
    addr = re.sub(r'\s+', ' ', addr)      # remove double spaces
    addr = re.sub(r'\bstreet\b', 'st', addr)  # optional abbreviation normalization
    addr = re.sub(r'\broad\b', 'rd', addr)
    return addr

combined["CleanAddress"] = combined["Address"].apply(clean_address)

# Drop duplicates based on cleaned version
unique_addresses = combined.drop_duplicates(subset=["CleanAddress"], keep="first")

# (Optional) Remove the helper column
unique_addresses = unique_addresses.drop(columns=["CleanAddress"])

# Save the result
unique_addresses.to_excel("Oct_Addresses.xlsx", index=False)

print(f"✅ Saved {len(unique_addresses)} unique addresses to Oct_Addresses.xlsx")

