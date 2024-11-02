import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def expand_comma_separated_columns(excel_file, separator=","):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(excel_file)
    
    # Create a new DataFrame to store the results
    result_df = df.copy()  # Start with the original data
    
    # Process only columns with comma-separated values
    for column in df.columns:
        # Check if any cell in the column has a comma (indicating it is comma-separated)
        if df[column].dropna().astype(str).str.contains(separator).any():
            # Collect unique values across all cells in this column
            unique_values = set()
            for cell in df[column].dropna().astype(str):
                unique_values.update(cell.split(separator))
            
            # Clean up whitespace in unique values
            unique_values = {value.strip() for value in unique_values}
            
            # For each unique value, create a new column with the format "<original_column>_<value>"
            for value in unique_values:
                # Column name includes original column and unique value
                new_column_name = f"{column}_{value.strip()}"
                
                # Fill the new column with 1 if the value is present in the cell, else 0
                result_df[new_column_name] = df[column].apply(
                    lambda cell: int(value.strip() in cell.split(separator)) if pd.notna(cell) else 0
                )
    
    return result_df

def process_file():
    # Ask user for file path
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_path:
        return  # User canceled the file dialog
    
    # Get the separator input from the user
    separator = separator_entry.get()
    if not separator:
        messagebox.showerror("Error", "Please enter a separator.")
        return
    
    try:
        # Process the file and create the expanded DataFrame
        expanded_df = expand_comma_separated_columns(file_path, separator)
        
        # Save the result to a new Excel file in the same directory
        output_file = os.path.join(os.path.dirname(file_path), "expanded_output.xlsx")
        expanded_df.to_excel(output_file, index=False)
        
        # Show success message
        messagebox.showinfo("Success", f"Expanded Excel file created as '{output_file}'")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Set up the GUI
root = tk.Tk()
root.title("Excel Column Expander")
root.geometry("400x200")

# Instruction label
instruction_label = tk.Label(root, text="Choose an Excel file and enter a separator:")
instruction_label.pack(pady=10)

# Separator entry
separator_label = tk.Label(root, text="Separator:")
separator_label.pack()
separator_entry = tk.Entry(root, width=10)
separator_entry.insert(0, ",")  # Default separator
separator_entry.pack(pady=5)

# Process button
process_button = tk.Button(root, text="Select File and Process", command=process_file)
process_button.pack(pady=20)

# Run the main event loop
root.mainloop()
