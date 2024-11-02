import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def expand_comma_separated_columns(excel_file, separator=","):
    df = pd.read_excel(excel_file)
    result_df = df.copy()
    
    for column in df.columns:
        # Check if ANY cell in the column contains the separator
        if df[column].dropna().astype(str).str.contains(separator).any():
            # Collect all unique values, handling both separated and non-separated cells
            unique_values = set()
            for cell in df[column].dropna().astype(str):
                # If the cell contains the separator, split it
                if separator in cell:
                    unique_values.update(value.strip() for value in cell.split(separator))
                else:
                    # If no separator, add the whole cell value
                    unique_values.add(cell.strip())
            
            # Create binary columns for each unique value
            for value in unique_values:
                new_column_name = f"{column}_{value.strip()}"
                
                # Modified lambda function to handle both cases
                result_df[new_column_name] = df[column].apply(
                    lambda cell: 1 if pd.notna(cell) and (
                        value.strip() == cell.strip() or  # Exact match for non-separated values
                        (separator in cell and value.strip() in [v.strip() for v in cell.split(separator)])  # Match in separated values
                    ) else 0
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

