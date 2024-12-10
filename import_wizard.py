import pandas as pd
import pyodbc
from tkinter import Tk, filedialog, simpledialog, messagebox

def import_excel_to_sql():
    # Open file dialog to select Excel file
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        messagebox.showerror("Error", "No file selected.")
        return

    # Prompt user for additional inputs
    sheet_name = simpledialog.askstring("Sheet Name", "Enter the sheet name (default: Sheet1):", initialvalue="Sheet1")
    table_name = simpledialog.askstring("Table Name", "Enter the table name:")
    server = simpledialog.askstring("Server Name", "Enter SQL Server name:")
    database = simpledialog.askstring("Database Name", "Enter database name:")
    username = simpledialog.askstring("Username", "Enter SQL Server username (optional):")
    password = simpledialog.askstring("Password", "Enter SQL Server password (optional):", show="*")

    # Validate inputs
    if not all([file_path, table_name, server, database]):
        messagebox.showerror("Error", "File, Table, Server, and Database fields are required.")
        return

    # SQL Server connection configuration (username and password are optional)
    if username and password:
        connection_string = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}"
    else:
        connection_string = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database}"

    try:
        conn = pyodbc.connect(connection_string)
        cursor = conn.cursor()
    except Exception as e:
        messagebox.showerror("Error", f"Database connection failed: {e}")
        return

    # Load the Excel file into a DataFrame
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        messagebox.showerror("Error", f"Error reading Excel file: {e}")
        return

    # Create table
    columns = df.columns
    column_definitions = ", ".join([f"[{col}] NVARCHAR(MAX)" for col in columns])
    create_table_query = f"CREATE TABLE {table_name} ({column_definitions});"
    try:
        cursor.execute(create_table_query)
    except Exception as e:
        messagebox.showerror("Error", f"Error creating table: {e}")
        conn.rollback()
        return

    # Insert data
    for index, row in df.iterrows():
        placeholders = ", ".join(["?"] * len(row))
        insert_query = f"INSERT INTO {table_name} ({', '.join([f'[{col}]' for col in columns])}) VALUES ({placeholders})"
        try:
            cursor.execute(insert_query, tuple(row))
        except Exception as e:
            messagebox.showerror("Error", f"Error inserting row {index}: {e}")
            conn.rollback()

    # Commit and close connection
    conn.commit()
    conn.close()
    messagebox.showinfo("Success", "Data imported successfully.")

# Run the GUI
if __name__ == "__main__":
    root = Tk()
    root.withdraw()  # Hide the root window
    import_excel_to_sql()
