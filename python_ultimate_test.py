import datetime
import random
import string
import tkinter as tk
from tkinter import ttk
import pandas as pd
from tkinter import filedialog
from tkinter import messagebox
import csv
import time
from tkinter import simpledialog

# Function to show the main application after the loading screen
def show_application():
    loading_screen.destroy()
    root.deiconify()

# Function to show a success message
def show_success_message(message):
    messagebox.showinfo('Success', message)

# Function to show an error message
def show_error_message(message):
    messagebox.showerror('Error', message)

# Load the CSV file into a DataFrame
file_path = 'C:\\Users\\micha\\Documents\\test_table.csv'
df = pd.read_csv(file_path)

# Create the main window
root = tk.Tk()
root.title('Financial Data')

# Set the application width and padding
width = int(root.winfo_screenwidth() * 0.8)
padding = 20
root.geometry(f'{width}x500')
root.configure(padx=padding, pady=padding)
root.configure(bg='#F4F7F9')

# Create a frame for the filters
filter_frame = tk.Frame(root, bg='#F4F7F9')
filter_frame.pack(pady=10)

# Get the list of column headers and their order
columns_order = df.columns.tolist()

# Create filters for each field
filters = {}
for i, column in enumerate(columns_order):
    unique_values = df[column].unique()
    filter_label = tk.Label(filter_frame, text=f'{column}:', font=('Arial', 10), bg='#F4F7F9')
    filter_label.grid(row=0, column=i, padx=5, pady=5, sticky='n')
    filter_label.configure(anchor='center')
    filter_var = tk.StringVar(value='All')
    if pd.api.types.is_numeric_dtype(df[column]):
        filter_entry = ttk.Entry(filter_frame, textvariable=filter_var, font=('Arial', 10))
        filter_entry.grid(row=1, column=i, padx=5, pady=5)
        filters[column] = filter_var
    else:
        filter_dropdown = ttk.Combobox(filter_frame, textvariable=filter_var, values=['All'] + list(unique_values),
            font=('Arial', 10))
        filter_dropdown.grid(row=1, column=i, padx=5, pady=5)
        filters[column] = filter_var

# Create a treeview to display the data
columns = columns_order
treeview = ttk.Treeview(root, columns=columns, show='headings', style='Custom.Treeview')
for col in columns:
    treeview.heading(col, text=col, anchor='center')
    treeview.column(col, width=int((width - 2 * padding) / len(columns)), anchor='center')
treeview.pack(fill='both', expand=True)

# Configure treeview style
style = ttk.Style()
style.configure('Custom.Treeview', background='#FFFFFF', foreground='#454545')
style.map('Custom.Treeview', background=[('selected', '#C6E6FF')])

# Style for the treeview headers
style.configure('Custom.Treeview.Heading', font=('Arial', 10, 'bold'), foreground='#454545', background='#EDEDED')

def get_sort_icon(sort_order):
    if sort_order == 'asc':
        return '\uf0de'  # Font Awesome icon for ascending sort
    elif sort_order == 'desc':
        return '\uf0dd'  # Font Awesome icon for descending sort
    else:
        return ''

# Function to populate the treeview with data based on filters
def populate_treeview():
    filtered_df = df.copy()
    for column, var in filters.items():
        selected_value = var.get()
        if selected_value != 'All':
            if pd.api.types.is_numeric_dtype(df[column]):
                try:
                    selected_value = float(selected_value)
                    filtered_df = filtered_df[filtered_df[column] == selected_value]
                except ValueError:
                    continue
            else:
                filtered_df = filtered_df[filtered_df[column] == selected_value]
    treeview.delete(*treeview.get_children())
    for row in filtered_df.itertuples(index=False):
        treeview.insert('', 'end', values=row)

# Function to export data to a CSV file
def export_data():
    filtered_df = df.copy()
    for column, var in filters.items():
        selected_value = var.get()
        if selected_value != 'All':
            if pd.api.types.is_numeric_dtype(df[column]):
                try:
                    selected_value = float(selected_value)
                    filtered_df = filtered_df[filtered_df[column] == selected_value]
                except ValueError:
                    continue
            else:
                filtered_df = filtered_df[filtered_df[column] == selected_value]
    export_file_path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV Files', '*.csv')])
    if export_file_path:
        try:
            filtered_df.to_csv(export_file_path, index=False)
            show_success_message(f'Data exported to: {export_file_path}')
        except Exception as e:
            show_error_message(f'Error occurred while exporting data: {str(e)}')

# Function to append data from a CSV file
def append_data():
    global df
    file_path = filedialog.askopenfilename(filetypes=[('CSV Files', '*.csv')])
    if file_path:
        try:
            new_data = pd.read_csv(file_path)
            
            # Add date column with current date
            new_data['Date'] = datetime.datetime.now().strftime('%Y-%m-%d')
            
            # Generate incremental numbers
            existing_rows = len(df)
            new_data['IncrementalNumbers'] = range(existing_rows + 1, existing_rows + 1 + len(new_data))
            
            # Fill missing values with random data
            for column in df.columns:
                if column not in new_data.columns:
                    if pd.api.types.is_numeric_dtype(df[column]):
                        new_data[column] = random.randint(0, 100)
                    elif pd.api.types.is_string_dtype(df[column]):
                        random_string = ''.join(random.choices(string.ascii_letters + string.digits, k=len(new_data)))
                        new_data[column] = random_string
                    else:
                        new_data[column] = random.choice([True, False])
            
            df = pd.concat([df, new_data], ignore_index=True)
            populate_treeview()
            show_success_message('Data appended successfully.')
        except Exception as e:
            show_error_message(f'Error occurred while appending data: {str(e)}')

# Function to delete rows
def delete_rows():
    selected_items = treeview.selection()
    if not selected_items:
        show_error_message('No rows selected. Please select rows to delete.')
        return
    # Prompt user to enter a reason for deletion
    reason = simpledialog.askstring('Reason for Deletion', 'Enter a reason for deleting the rows:')
    if reason is None:
        return
    confirmation = messagebox.askyesno('Confirmation', 'Are you sure you want to delete the selected rows?')
    if confirmation:
        deleted_rows = []
        for item in selected_items:
            row = treeview.item(item, 'values')
            deleted_rows.append(row)
            treeview.delete(item)

        # Add a new column for comments
        deleted_rows_with_comments = []
        for row in deleted_rows:
            row_with_comments = tuple(row) + (reason,)
            deleted_rows_with_comments.append(row_with_comments)

        show_success_message('Rows deleted successfully.')

        # Save deleted rows to a CSV file with current date and time
        current_datetime = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        deleted_file_path = f'C:\\Users\\micha\\Documents\\deleted_rows_{current_datetime}.csv'
        with open(deleted_file_path, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(columns + ['Reason'])  # Write the column headers with the new 'Reason' column
            writer.writerows(deleted_rows_with_comments)  # Write the deleted rows with comments

sort_orders = {column: 'none' for column in columns}

def on_header_click(col):
    sort_order = sort_orders[col]
    if pd.api.types.is_numeric_dtype(df[col]):
        if sort_order == 'none' or sort_order == 'asc':
            df.sort_values(col, ascending=False, inplace=True)
            sort_orders[col] = 'desc'
        else:
            df.sort_values(col, ascending=True, inplace=True)
            sort_orders[col] = 'asc'
    else:
        if sort_order == 'none' or sort_order == 'asc':
            df.sort_values(col, ascending=False, inplace=True, key=lambda x: x.str.lower())
            sort_orders[col] = 'desc'
        else:
            df.sort_values(col, ascending=True, inplace=True, key=lambda x: x.str.lower())
            sort_orders[col] = 'asc'
    populate_treeview()

    # Update column headers with sort icons
    for col in columns:
        sort_icon = get_sort_icon(sort_orders[col])
        treeview.heading(col, text=f'{col} {sort_icon}', anchor='center', command=lambda c=col: on_header_click(c))

# Create the treeview with initial column headers
for col in columns:
    sort_icon = get_sort_icon(sort_orders[col])
    treeview.heading(col, text=f'{col} {sort_icon}', anchor='center', command=lambda c=col: on_header_click(c))

# Simulating loading the data
time.sleep(3)  # Wait for 3 seconds to simulate loading

# Get today's date and time
current_datetime = datetime.datetime.now()
date_string = current_datetime.strftime('%Y-%m-%d')
time_string = current_datetime.strftime('%H:%M:%S')

# Create the loading screen
loading_screen = tk.Tk()
loading_screen.title('Loading...')
loading_screen.geometry('400x200')
loading_screen.configure(bg='#F4F7F9')

# Center the loading screen window
screen_width = loading_screen.winfo_screenwidth()
screen_height = loading_screen.winfo_screenheight()
x = int((screen_width - loading_screen.winfo_reqwidth()) / 2)
y = int((screen_height - loading_screen.winfo_reqheight()) / 2)
loading_screen.geometry(f'+{x}+{y}')

# Create a frame for the loading screen
loading_frame = tk.Frame(loading_screen, bg='#F4F7F9')
loading_frame.pack(expand=True)

# Create a label for the loading message
label_message = tk.Label(loading_frame, text='Loading Data...', font=('Arial', 16), bg='#F4F7F9')
label_message.pack(pady=40)

# Call the function to show the main application after 2 seconds
loading_screen.after(5000, show_application)  # Display loading screen for 5 seconds (5000 milliseconds)

# Populate the treeview with data based on filters
populate_treeview()

# Bind filter changes to populate the treeview dynamically
for var in filters.values():
    var.trace('w', lambda *args: populate_treeview())

# Create the Export Data button
export_button = ttk.Button(root, text='Export Data', command=export_data)
export_button.pack(pady=10)

# Create the Append Data button
append_button = ttk.Button(root, text='Append Data', command=append_data)
append_button.pack(pady=10)

# Create the Delete Rows button
delete_button = ttk.Button(root, text='Delete Rows', command=delete_rows)
delete_button.pack(pady=10)

# Start the GUI event loop
root.withdraw()  # Hide the main application initially
loading_screen.mainloop()
