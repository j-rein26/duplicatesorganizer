import pandas as pd

# 1️⃣ Read Excel files
df1 = pd.read_excel("NewHomeBuyersOct.xlsx")
df2 = pd.read_excel("NewMoversOct.xlsx")

# 2️⃣ Print column names
print("File 1 columns:", df1.columns.tolist())
print("File 2 columns:", df2.columns.tolist())

# 3️⃣ Print number of rows
print(f"File 1 rows: {len(df1)}")
print(f"File 2 rows: {len(df2)}")

# 4️⃣ Preview first 20 addresses from each file
if "Address" in df1.columns:
    print("File 1 addresses (first 20):")
    print(df1["Address"].head(20))
else:
    print("File 1 does NOT have a column named 'Address'")

if "Address" in df2.columns:
    print("File 2 addresses (first 20):")
    print(df2["Address"].head(20))
else:
    print("File 2 does NOT have a column named 'Address'")

# 5️⃣ Combine the files (for inspection only)
combined = pd.concat([df1, df2], ignore_index=True)
print(f"Combined rows before removing duplicates: {len(combined)}")

# 6️⃣ Preview first 20 addresses in combined DataFrame
if "Address" in combined.columns:
    print("Combined addresses (first 20):")
    print(combined["Address"].head(20))
