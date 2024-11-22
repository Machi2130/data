import pandas as pd
import tkinter as tk
from tkinter import filedialog
import difflib

def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    return file_path

def compare_excel_sheets():
    # Get file paths through file dialog
    print("Select first Excel file:")
    file1 = select_file()
    print("Select second Excel file:")
    file2 = select_file()
    
    # Read Excel files
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    
    # Convert dataframes to string format for comparison
    df1_str = df1.astype(str)
    df2_str = df2.astype(str)
    
    # Initialize differences dictionary
    differences = {
        'row': [],
        'column': [],
        'file1_value': [],
        'file2_value': []
    }
    
    # Compare each cell
    for col in df1.columns:
        for idx in range(len(df1)):
            if df1_str.iloc[idx][col] != df2_str.iloc[idx][col]:
                differences['row'].append(idx + 2)  # Adding 2 because Excel rows start from 1 and header is row 1
                differences['column'].append(col)
                differences['file1_value'].append(df1_str.iloc[idx][col])
                differences['file2_value'].append(df2_str.iloc[idx][col])
    
    # Create differences DataFrame
    diff_df = pd.DataFrame(differences)
    
    # Save differences to Excel file
    output_file = 'differences_report.xlsx'
    diff_df.to_excel(output_file, index=False)
    print(f"Differences have been saved to {output_file}")
    
    # Display differences in console
    if len(diff_df) > 0:
        print("\nDifferences found:")
        print(diff_df)
    else:
        print("\nNo differences found between the Excel sheets.")

if __name__ == "__main__":
    compare_excel_sheets()
