from logging import root
import smartsheet.sheets
import warnings
warnings.simplefilter(action='ignore',category=FutureWarning)
import pandas as pd
import os
from datetime import datetime
import tkinter as tk
from tkinter import ttk
#Default Values
Default_Access_Token = 'YOUR SMARTSHEET ACCESS TOKEN'
Default_Sheet_IDs = ['YOUR SMARTSHEET ID 1', 'YOUR SMARTSHEET ID 2']

def log_message(log_text, message, tag="normal"): 
    log_text.insert(tk.END, message + "\n", tag)
    log_text.see(tk.END) 
    root.update_idletasks()

#Create Smartsheet CLient
def download_sheets(Default_Access_Token, Default_Sheet_IDs, progress_var,log_text):

    ss_client = smartsheet.Smartsheet(Default_Access_Token)
    log_message(log_text, "Connected To Smartsheet", "success")

    download_time = datetime.now().strftime("%d-%m-%Y %H:%M")
    total_sheets = len(Default_Sheet_IDs)

    #Download Each Sheet In CSV
    for idx, sheet_id in enumerate(Default_Sheet_IDs):

        #Get Sheet Data
        try:
            response = ss_client.Sheets.get_sheet(sheet_id)
            log_message(log_text, f"{response.name} Data Collected", "success")

            #Check If response is an instance or Error object
            if isinstance(response, smartsheet.models.Error):
                log_message(log_text, f"Error: {response.message}", "error")
                continue

            sheet_name = response.name

            #Define File path
            file_path = os.path.join(os.getcwd(),f"{sheet_name}.xlsx")

            #Create Data Yousing Pandas
            data = []
            for row in response.rows:
                row_data = [] #{col.title: '' for col in response.columns} # Initialise row with empty strings
                for cell in row.cells:
                    #column_title = next((col.title for col in response.columns if col.id == cell.column_id), None)
                    cell_value = cell.display_value if cell.display_value else cell.value
                    if cell_value is not None:
                        try:
                            cell_value = pd.to_numeric(cell_value,errors='ignore')
                        except(ValueError,TypeError):
                            pass
                    row_data.append(cell_value)
                data.append([download_time] + row_data)

            columns = (["Download Time"] + [column.title for column in response.columns])
            
            df = pd.DataFrame(data,columns=columns)
                        
            #Format Dates
            for column in df.columns:
                if df[column].dtype == 'object':
                    try:
                        df[column] = pd.to_datetime(df[column], format= '%d-%m-%Y')
                    except(ValueError,TypeError):
                        pass
            
            log_message(log_text, "Data Combined", "success")
            
            df.to_excel(file_path,index=False)
            log_message(log_text, f"{sheet_name} Downloaded As {file_path}", "success")

        except Exception as e:
            log_message(log_text, f"Failed to Download {sheet_name}: {e}", "error")
        
        progress_var.set((idx+1)/total_sheets*100)
        root.update_idletasks()
          
    log_message(log_text,"All sheets have been downloaded successfully", "success")

#Tkinter GUI Setup
root = tk.Tk()
root.title("SmartSheet Downloader")
root.geometry("600x300")

progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
progress_bar.pack(pady=20, padx=20, fill=tk.X)

log_text = tk.Text(root,height=10)
log_text.pack(pady=5,padx=5, fill=tk.BOTH, expand=True)

# Define text tags for styling 
log_text.tag_configure("success", foreground="green", font=("TkDefaultFont", 10, "bold")) 
log_text.tag_configure("error", foreground="red", font=("TkDefaultFont", 10, "bold"))

#Run The Downloader
root.after(100, lambda:download_sheets(Default_Access_Token, Default_Sheet_IDs, progress_var, log_text))
root.mainloop()
