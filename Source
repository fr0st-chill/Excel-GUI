import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import subprocess
import sys

def check_dependencies():
    # Read the requirements.txt file and install the dependencies
    try:
        with open('requirements.txt', 'r') as f:
            dependencies = f.read().splitlines()
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', *dependencies])
    except FileNotFoundError:
        pass
    
def reset_error_label():
    error_label.config(text='')

def browse_file(entry_widget):
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, file_path)

def browse_file_save(entry_widget):
    file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, file_path)

def match_data():
    file_path1 = entry_file1.get()
    file_path2 = entry_file2.get()
    target_column1 = entry_target1.get()
    target_column2 = entry_target2.get()
    output_file_path = entry_output.get()

    # Check if target columns are specified
    if not target_column1 or not target_column2:
        error_label.config(text="Please specify target columns.")
        window.after(3000, reset_error_label)  # Schedule resetting error label after 3 seconds
        return

    try:
        # Read the CSV or XLSX files
        df1 = pd.read_csv(file_path1) if file_path1.endswith('.csv') else pd.read_excel(file_path1)
        df2 = pd.read_csv(file_path2) if file_path2.endswith('.csv') else pd.read_excel(file_path2)

        # Check if target columns exist
        if target_column1 not in df1.columns or target_column2 not in df2.columns:
            error_label.config(text="Target column does not exist.")
            return

        # Perform data matching based on target columns (string matching)
        merged_df = pd.merge(df1, df2, on=[target_column1, target_column2])

        # Save the output file
        merged_df.to_csv(output_file_path, index=False)

        # Show success message
        messagebox.showinfo("Success", "File created successfully!")
        
        # Reset error label
        error_label.config(text="")

    except FileNotFoundError:
        error_label.config(text="File not found.")
        window.after(3000, reset_error_label)  # Schedule resetting error label after 3 seconds


# Create the GUI window
window = tk.Tk()
window.title('Excel Sheet Comparison')

# Create file input fields
label_file1 = tk.Label(window, text='Excel File 1:')
label_file1.grid(row=0, column=0, padx=10, pady=10)
entry_file1 = tk.Entry(window, width=50)
entry_file1.grid(row=0, column=1, padx=10, pady=10)
browse_button1 = tk.Button(window, text='Browse', command=lambda: browse_file(entry_file1))
browse_button1.grid(row=0, column=2, padx=10, pady=10)

label_file2 = tk.Label(window, text='Excel File 2:')
label_file2.grid(row=1, column=0, padx=10, pady=10)
entry_file2 = tk.Entry(window, width=50)
entry_file2.grid(row=1, column=1, padx=10, pady=10)
browse_button2 = tk.Button(window, text='Browse', command=lambda: browse_file(entry_file2))
browse_button2.grid(row=1, column=2, padx=10, pady=10)

# Create target column input fields
label_target1 = tk.Label(window, text='Target Column 1:')
label_target1.grid(row=2, column=0, padx=10, pady=10)
entry_target1 = tk.Entry(window, width=50)
entry_target1.grid(row=2, column=1, padx=10, pady=10)

label_target2 = tk.Label(window, text='Target Column 2:')
label_target2.grid(row=3, column=0, padx=10, pady=10)
entry_target2 = tk.Entry(window, width=50)
entry_target2.grid(row=3, column=1, padx=10, pady=10)

# Create output file input field
label_output = tk.Label(window, text='Output CSV File:')
label_output.grid(row=4, column=0, padx=10, pady=10)
entry_output = tk.Entry(window, width=50)
entry_output.grid(row=4, column=1, padx=10, pady=10)
browse_button3 = tk.Button(window, text='Browse', command=lambda: browse_file_save(entry_output))
browse_button3.grid(row=4, column=2, padx=10, pady=10)

# Create run button
run_button = tk.Button(window, text='Run', command=match_data)
run_button.grid(row=5, column=1, padx=10, pady=10)

# Create error label
error_label = tk.Label(window, text='', fg='red')
error_label.grid(row=6, column=0, columnspan=3, padx=10, pady=10)

window.mainloop()
