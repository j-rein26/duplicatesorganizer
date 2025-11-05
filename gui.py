import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

def standardize_columns(df, file_label):
    """
    Ensure the dataframe has columns: First, Last, Address, City, State, Zip.
    If FullName exists, split it into First and Last.
    """
    expected_cols = ["First", "Last", "Address", "City", "State", "Zip"]

    # Handle FullName column
    if "FullName" in df.columns:
        df[["First", "Last"]] = df["FullName"].astype(str).str.split(" ", 1, expand=True)

    # Add missing columns with empty strings
    for col in expected_cols:
        if col not in df.columns:
            df[col] = ""

    # Keep only expected columns in correct order
    df = df[expected_cols]
    return df

def combine_addresses():
    # Select first file
    file1_path = filedialog.askopenfilename(title="Select First Excel File", filetypes=[("Excel files", "*.xlsx")])
    if not file1_path:
        return

    # Select second file
    file2_path = filedialog.askopenfilename(title="Select Second Excel File", filetypes=[("Excel files", "*.xlsx")])
    if not file2_path:
        return

    try:
        # Read files
        df1 = pd.read_excel(file1_path)
        df2 = pd.read_excel(file2_path)

        # Standardize columns
        df1 = standardize_columns(df1, "File 1")
        df2 = standardize_columns(df2, "File 2")

        # Check for missing critical column (Address)
        for i, df in enumerate([df1, df2], start=1):
            if df["Address"].isnull().all():
                messagebox.showerror("Error", f"File {i} does not have an Address column or it's empty.")
                return

        # Combine
        combined = pd.concat([df1, df2], ignore_index=True)

        # Clean addresses for deduplication
        combined["Address"] = combined["Address"].astype(str).str.strip().str.lower()

        # Remove duplicates based on Address
        unique_addresses = combined.drop_duplicates(subset=["Address"], keep="first")

        # Restore formatting
        for col in ["First", "Last", "City", "State"]:
            unique_addresses[col] = unique_addresses[col].str.title()
        unique_addresses["Address"] = unique_addresses["Address"].str.title()
        unique_addresses["Zip"] = unique_addresses["Zip"].astype(str).str.strip()

        # --- Preview ---
        total_rows = len(combined)
        unique_rows = len(unique_addresses)
        duplicates_count = total_rows - unique_rows

        preview_win = tk.Toplevel(root)
        preview_win.title("Preview & Summary")
        preview_win.geometry("600x450")

        summary_text = f"File 1 rows: {len(df1)}\n" \
                       f"File 2 rows: {len(df2)}\n" \
                       f"Total combined rows: {total_rows}\n" \
                       f"Duplicates found: {duplicates_count}\n" \
                       f"Unique addresses: {unique_rows}\n\n" \
                       f"First 20 rows:\n" \
                       f"{unique_addresses[['First','Last','Address','City','State','Zip']].head(20).to_string(index=False)}"

        txt = scrolledtext.ScrolledText(preview_win, wrap=tk.WORD)
        txt.insert(tk.END, summary_text)
        txt.configure(state="disabled")
        txt.pack(expand=True, fill='both', padx=10, pady=10)

        # --- Save Button ---
        def save_file():
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                     initialfile="Oct_Addresses.xlsx",
                                                     filetypes=[("Excel files", "*.xlsx")])
            if save_path:
                unique_addresses.to_excel(save_path, index=False)
                messagebox.showinfo("Success", f"✅ Saved {unique_rows} unique addresses to {save_path}")
                preview_win.destroy()

        save_btn = tk.Button(preview_win, text="Save Deduplicated File", command=save_file, width=30, height=2)
        save_btn.pack(pady=10)

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")


# --- Tkinter GUI ---
root = tk.Tk()
root.title("Address Combiner - Robust Version")
root.geometry("500x220")

label = tk.Label(root, text="Combine two Excel files with columns:\nFirst, Last, Address, City, State, Zip\nHandles missing/misnamed columns and duplicates",
                 wraplength=450, justify="center")
label.pack(pady=10)

button = tk.Button(root, text="Select Files & Run", command=combine_addresses, width=28, height=2)
button.pack(pady=20)

root.mainloop()


