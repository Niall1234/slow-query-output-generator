import os
import re
import string
import openpyxl
import csv
import pymysql
from tkinter import *
from tkinter import filedialog
from dotenv import load_dotenv


# list of available DBs
databases = {
    "smartidplus": "db-sq-prod-smartidplus.eu.lnrm.net",
    "crystal": "db-sq-prod-crystal.eu.lnrm.net",
    "smartcleanse portal": "db-sq-prod-smartcleanse-bureau.eu.lnrm.net"
}

# functions
def execute(database, query, output_format):
    if not validate_query(query):
        error_output("Statement must start with SELECT'")
        return
    
    try:
        connection = pymysql.connect(
            host=databases[database],
            user=os.getenv("USERNAME"),
            password=os.getenv("PASSWORD"),
            db=database,
            charset="utf8mb4",
            cursorclass=pymysql.cursors.DictCursor
        )
    except Exception as exc:
        error_output(exc)
        return

    try:
        columns = get_columns(connection, query)
        result = run_query(connection, query)
    except Exception as exc:
        error_output(exc) 

    print(result)

    # chosen file naming convention
    filename = filename_entry.get() 

    if output_format == "excel":
        output_to_xl(result, columns, filename)
    elif output_format == "csv":
        output_to_csv(result, columns, filename)


# Parse SQL query for output columns
def get_columns(connection, query):
    with connection.cursor() as cursor:
        cursor.execute(query)
        columns = list(cursor.fetchone().keys())
        return columns
    
# Execute and return main query
def run_query(connection, query):
    with connection.cursor() as cursor:
        cursor.execute(query)
        result = cursor.fetchall()
        return result

    
# check that SQL query contains a proper statement
def validate_query(query):
    return query.lower().startswith("select")
    

# TO DO: function to create CSV file from SQL query
def output_to_csv(result, columns, filename):
    if f"{filename}.csv" in os.listdir():
        count = 1
        while f"{filename}-{count}.csv" in os.listdir():
            count += 1
        csv_file = f"{filename}-{count}.csv"
    else:
        csv_file = f"{filename}.csv"
    
    with open(csv_file, "w", newline="") as open_file:
        csv_writer = csv.writer(open_file, delimiter=",", quotechar="'")
        csv_writer.writerow(columns)
        for i in range(len(result)):
            csv_writer.writerow(list(result[i].values()))


# function to create Excel file from SQL query
def output_to_xl(result, columns, filename):
    excel_columns = list(string.ascii_uppercase) + list("A" + char for char in string.ascii_uppercase) + list("B" + char for char in string.ascii_uppercase)
    wb = openpyxl.Workbook()
    sheet = wb.active

    for pos, column in enumerate(columns):
        sheet[f"{excel_columns[pos]}1"].value = column
    
        for j in range(len(result)):
            sheet[f"{excel_columns[pos]}{j + 2}"].value = result[j][sheet[f"{excel_columns[pos]}1"].value]


    if f"{filename}.xlsx" in os.listdir():
        count = 1
        while f"{filename}-{count}.xlsx" in os.listdir():
            count += 1
        wb.save(f"{filename}-{count}.xlsx")
    else:
        wb.save(f"{filename}.xlsx")


# TO DO: Add option to upload SQL file
def upload_file():
    file_path = filedialog.askopenfilename(title="select file", filetypes=[("sql files", "*.sql"), ("text files", "*.txt")])
    if file_path:
        selected_file_label.config(text=f"selected file: {file_path}")

    try:
        with open(file_path, "r") as file:
            contents = file.read()
            sql_field.delete(1.0, END)
            sql_field.insert(END, contents)
    except Exception as exc:
        error_output(exc)
             
# TO DO: error handling in GUI
def error_output(error):
    errors_field.delete(1.0, END)
    errors_field.insert(1.0, error)

# Building TK interface
window = Tk()
window.title("SQL Slow Query")
master_frame = Frame(window)
master_frame.pack(expand = True, fill = BOTH)

sql_frame = Frame(master=master_frame, padx=5, pady=5)
sql_frame.pack(side = LEFT, expand = True)

sql_field = Text(master=sql_frame, width=50, height=20)
sql_field.pack(side = LEFT)
sql_field.insert(1.0, "Type your SQL here...")

button_frame = Frame(master=master_frame, padx=5, pady=5)
button_frame.pack(side = RIGHT, expand = True)

filename_label = Label(master=button_frame, text="Enter filename for output")
filename_label.pack(expand = True)

filename_entry = Entry(master=button_frame, width=20)
filename_entry.pack(expand = True)

default_value = StringVar(master_frame)
default_value.set("Select a database from the list")
db_selector = OptionMenu( button_frame, default_value, *list(databases.keys()))
db_selector.pack(expand = True)

xl_button = Button(master=button_frame, width=20, height=2, text="Download XLSX", command= lambda: execute(default_value.get(), sql_field.get(1.0, END), "excel"))
xl_button.pack(expand = True)

csv_button = Button(master=button_frame, width=20, height=2, text="Download CSV", command= lambda: execute(default_value.get(), sql_field.get(1.0, END), "csv"))
csv_button.pack(expand = True)

upload_button = Button(master=button_frame, width=20, height=2, text="Upload SQL", command=upload_file)
upload_button.pack(expand = True)
selected_file_label = Label(master=button_frame, height=1)
selected_file_label.pack(expand = True, fill = BOTH)

errors_field = Text(master=button_frame, width=30, height=7, background="black", foreground="yellow")
errors_field.insert(1.0, "Errors will appear here...")
errors_field.pack(expand = True)

window.mainloop()


