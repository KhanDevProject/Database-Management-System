#############################################       Import or install all the necessary modules ## Otherwise the code won't work ##########################################################################################

import contextlib, os, tempfile, openpyxl, pandas as pd, win32print, tkinter as tk, tkinter.ttk as ttk 
from io import BytesIO
from tkinter import *
from tkinter import Button, Entry, Label, filedialog, messagebox
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

#########################################################################################################################################################################################################
# Define a class called database_management that inherits from the Frame class
class database_management(tk.Frame):
    def __init__(self, myapp):
        super().__init__(myapp)
        self.myapp = myapp # reference to the main application class, allows the class object to access myapp object directly
        # Create a title for myapp 
        myapp.title("Database Management System")
        myapp.resizable(False,False) # disable resizing of window
        self.data_loaded = False  # added variable to check if data has been loaded
        self.widgets_design() # call the widget_design method
        self.threshold_values() # call the threshold_values method
        
    def widgets_design(self): # define the widgets in the frame to avoid messy codes in the __init__ method
###########################       Creating Labels        ##########################################################
        self.label_title = tk.Label(self.myapp, text="CIS305 Database Management System", font=("Arial Bold", 10), bg="#0F7BAF", fg="white", highlightthickness=3, height=4, width=50) # create a label title for myapp 
        self.col_label = tk.Label(self.myapp, text="Please input the conditions in the following frame!", bg="#CCF760", fg="#1876A4", font=("Arial Bold", 11), relief=tk.RAISED, borderwidth=7, width=40) # create a label to show the conditions
        self.label_data_input = tk.Label(self.myapp, text="Data Input", relief=tk.SUNKEN, borderwidth=12, font=("Arial Bold", 10),height=2, width=9, fg="#02f2b6", bg="#808080") # create a label to show the path of the file 
        self.total_cost_label = tk.Label(self.myapp, text="Display Total Cost", font=("Arial Bold", 14), fg="#0895A1")

###########################       Creating Entry        ##########################################################
        self.file_location = tk.Entry(self.myapp, width=50) # create an entry for the path of the file
        self.entry_for_column = tk.Entry(self.myapp, width=18, font=('Arial', 25), highlightbackground='gray', highlightcolor='gray', highlightthickness=5, relief=tk.SUNKEN, borderwidth=8) # create an entry to show the conditions
        self.file_location2 = tk.Entry(self.myapp, width=50) # create an entry for the path of the file
        
################################       CREATING BUTTONS          #######################################################
        self.browse_button = tk.Button(self.myapp, text="Browse", command=self.file_browse, fg="white", bg='#808080', highlightcolor='#00BFFF', highlightthickness=1, width=8) # create a button to browse for a file
        self.button_for_search = tk.Button(self.myapp, text="Start", command=self.search_data_button, width=7, height=2, borderwidth=7, bg="#2EAA8A", fg="yellow", font=("Arial Bold", 11)) # create a button to search the data
        self.change_button = tk.Button(self.myapp, text="Change", command=self.change_popup_window, width=8, height=1, borderwidth=7, bg="red", fg="white", font=("Arial Bold", 11)) # create a button to change the path of the file
        self.button_for_more = tk.Button(self.myapp, text="More", width=9, height=1, borderwidth=5, bg="#18a2cc", fg="#7a090a", font=("Arial Bold", 11), command=self.more_rows) # create a button to add more rows
        self.button_for_save = tk.Button(self.myapp, text="Save", width=8, height=1, borderwidth=7, bg="#17CBE3", fg="white", font=("Arial Bold", 11), command=self.save_table_data) # create a button to save the data
        self.button_for_export = tk.Button(self.myapp, text="Export", width=8, height=1, borderwidth=7, bg="#4CC270", fg="white", font=("Arial Bold", 11)) # create a button to export
        self.browse_button2 = tk.Button(self.myapp, text="Browse", command=self.file_browse2, fg="white", bg='#808080', highlightcolor='#00BFFF', highlightthickness=1, width=8) # create a button to browse for a file
        self.button_for_level = tk.Button(self.myapp, text="Reset", width=6, height=2, borderwidth=7, bg="#1977C6", fg="white", font=("Arial Bold", 11), command=self.level_filter_data)

################################       BUTTONS PLACEMENT        #######################################################
        self.label_data_input.place(x=9, y=95) # place the label at the top of the window 
        self.total_cost_label.place(x=150, y=510) # place the label at the bottom of the window
        self.label_title.place(x=20, y=6) # place the label title at the top of the window
        self.change_button.place(x=970, y=320) # place the button at the bottom of the window
        self.button_for_more.place(x=70, y=320) # place the button at the bottom of the window
        self.btn_more_clicks = 0 # added variable to count the number of clicks on the more button
        self.button_for_save.place(x=970, y=380) # place the button at the bottom of the window
        self.button_for_export.place(x=970, y=440) # place the button at the bottom of the window
        self.button_for_export.config(command=self.button_export_print) # set the command of the export button
        self.file_location2.place(x=110, y=130) # place the entry at the bottom of the window
        self.browse_button2.place(x=420, y=127) # place the button at the bottom of the window
        self.button_for_level.place(x=970, y=190) # place the button at the bottom of the window
        self.file_location.place(x=110, y=100) # place the entry at the bottom of the window
        self.browse_button.place(x=420, y=98) # place the button at the bottom of the window
        self.col_label.place(x=510, y=10) # place the label at the bottom of the window
        self.entry_for_column.place(x=510, y=50) # place the entry at the bottom of the window
        self.button_for_search.place(x=870, y=52) # place the button at the bottom of the window

################################       START BUTTON        ###############################################################
        self.level = tk.IntVar(value=0) # added variable to store the level
        self.total_cost = tk.IntVar() # added variable to store the total cost
        self.frame_for_level = tk.Frame(self.myapp) # create a frame for the level
        self.frame_for_level.place(x=990, y=260) # place the frame at the top of the window
        tk.Radiobutton(self.frame_for_level, text="0", variable=self.level, value=0).grid(row=0, column=0, sticky='w') # create a radiobutton for level 0
        tk.Radiobutton(self.frame_for_level, text="1", variable=self.level, value=1).grid(row=1, column=0, sticky='w') # create a radiobutton for level 1

#################################   TREEVIEW WIDGET ######################################################
        self.display_info = tk.Listbox(self.myapp, bg='grey') # create a frame for the display
        self.display_info.place(x=500, y=160) 

        columns = ['ID', 'Total_Cost', 'Level', 'CustomerID'] 
        
        style = tk.ttk.Style()
        style.configure("Treeview.Heading", font=('Arial', 12, 'bold'), foreground="#333333")

        self.treeview = tk.ttk.Treeview(self.display_info, columns=columns, show="headings", height=18, style="Custom.Treeview")

        for col in columns:
            self.treeview.heading(col, text=col, anchor=tk.CENTER) 
            self.treeview.column(col, minwidth=0, width=110, stretch=tk.YES) 

        self.treeview.grid(row=0, column=0, sticky="nsew") 

        vertical_scrollbar = tk.Scrollbar(self.display_info, orient="vertical", command=self.treeview.yview)
        vertical_scrollbar.grid(row=0, column=1, sticky="ns")
        self.treeview.configure(yscrollcommand=vertical_scrollbar.set)

        style.configure("Custom.Treeview", background="#FFFFFF", foreground="#333333", rowheight=25)

        self.display_info.grid_rowconfigure(0, weight=1) 
        self.display_info.grid_columnconfigure(0, weight=1)

#################################   Total_Cost and radiobutton colors    ######################################################
        self.color_data = tk.StringVar(value='.') # added variable to store the color
        self.frame_for_color = tk.Frame(self.myapp) # create a frame for the color
        self.frame_for_color.place(x=140, y=540) # place the frame at the top of the window
        tk.Radiobutton(self.frame_for_color, text='Total cost (>=5000)', variable=self.color_data, value='red', fg='red', command=self.total_cost_color_updated).grid(row=0, column=0, sticky='w') # create a radiobutton for showing the color of the total cost based on threshold_values values which is RED
        tk.Radiobutton(self.frame_for_color, text='Total cost (1000~5000)', variable=self.color_data, value='green', fg='green', command=self.total_cost_color_updated).grid(row=1, column=0, sticky='w') # create a radiobutton for showing the color of the total cost based on threshold_values values which is GREEN 
        tk.Radiobutton(self.frame_for_color, text='Total cost (<=10000)', variable=self.color_data, value='grey', fg='grey', command=self.total_cost_color_updated).grid(row=2, column=0, sticky='w') # create a radiobutton for showing the color of the total cost based on threshold_values values which is GREY
    def threshold_values(self): # create a function to update the circles color based on the threshold_values value
        self.color_data.trace("w", lambda *args: self.update_circles_color(i,j)) # create a function to update the circles color based on the threshold_values value
        canvas = tk.Canvas(self.myapp, bg=self.myapp.cget('bg'), width=200, height=320) # create a canvas to draw the circles
        canvas.place(x=170, y=170) # place the canvas at the top of the window
        self.frame_for_threshold = tk.Frame(canvas, highlightthickness=0) # create a frame for the threshold
        self.frame_for_threshold.place(x=20, y=18, width=128, height=287) # place the frame at the top of the window
        self.labels_for_rows = list("ABCDEFGH") # create a list of labels for the rows
        self.labels_for_columns = ["", "1", "2", "3"] # create a list of labels for the columns
        self.store_circles = [] # create a list to store the circles
        self.circles_row_canvas = [] # create a list to store the circles on the canvas

        x1, y1, x2, y2 = 10, 10, 160, 312 # create the coordinates of the circles
        arc_len = 20 # create the arc length of the circles
        # create a circle and line into the canvas to draw the circles 
        canvas.create_arc(x1, y1, x1+arc_len*2, y1+arc_len*2, start=90, extent=90, style='arc', width=5, outline='#265588') 
        canvas.create_line(x1+arc_len, y1, x2-arc_len, y1, width=5, fill='#265588') 
        canvas.create_arc(x2-arc_len*2, y1, x2, y1+arc_len*2, start=0, extent=90, style='arc', width=5, outline='#265588') 
        canvas.create_line(x2, y1+arc_len, x2, y2-arc_len, width=5, fill='#265588') 
        canvas.create_arc(x2-arc_len*2, y2-arc_len*2, x2, y2, start=270, extent=90, style='arc', width=5, outline='#265588')
        canvas.create_line(x1+arc_len, y2, x2-arc_len, y2, width=5, fill='#265588')
        canvas.create_arc(x1, y2-arc_len*2, x1+arc_len*2, y2, start=180, extent=90, style='arc', width=5, outline='#265588')
        canvas.create_line(x1, y1+arc_len, x1, y2-arc_len, width=5, fill='#265588')
        
        for i, row_label in enumerate(self.labels_for_rows): # loop through the rows of the numbers 
            row = [] # create a list to store the row
            c = tk.Label(self.frame_for_threshold, text=row_label, font=("Arial", 12), padx=5, pady=5) # create a row label for the numbers
            c.grid(row=i+1, column=0, sticky="W") # Place it at the column 0 
            for j in range(len(self.labels_for_columns)-1): # loop through the columns of the numbers
                c = tk.Canvas(self.frame_for_threshold, width=30, height=30, highlightthickness=0) # create a canvas to draw the circles
                c.grid(row=i+1, column=j+1, padx=2, pady=1) # Place it at the column j+1
                oval = c.create_oval(5, 5, 25, 25, outline="#298da3", width=1) # create an oval
                row.append(oval) # add the oval to the row
            self.store_circles.append(row) # add the row to the list

        self.labels_for_columns = [" " if j == 0 else label_text for j, label_text in enumerate(self.labels_for_columns)] # create a list of labels for the columns
        [tk.Label(self.frame_for_threshold, text=label_text, font=("Arial", 12), padx=5, pady=5).grid(row=0, column=j) for j, label_text in enumerate(self.labels_for_columns)] # Place the labels at the columns 0 and 1
    def update_circles_color(self, i, j): # create a function to update the circles color based on the threshold_values value
        self.empty = [] # create a list to store the empty circles on the canvas
        circles = tk.Canvas(self.frame_for_threshold, width=30, height=30, highlightthickness=0) # create a canvas to draw the circles
        oval = circles.create_oval(5, 5, 25, 25, tags="oval") # create an oval
        circles.grid(row=i-6, column=2)  # Change the column to 2 
        self.empty.append(circles) # add the canvas to the list
        my_col = self.color_data.get() # get the threshold_values value
        if my_col == "red": # if the threshold_values value is RED
            circles.itemconfig(oval, fill="#ff0000") # change the oval my_col to red
        elif my_col == "green": # if the threshold_values value is GREEN
            circles.itemconfig(oval, fill="#00ff00") # change the oval color to green
        else: # if the threshold_values value is GREY
            circles.itemconfig(oval, fill="grey") # change the oval color to grey
    def button_export_print(self): # create a function to export and print the data
        # Read data from Excel sheet and append to PDF
        file_location1, file_location2 = self.file_location.get(), self.file_location2.get() # get the file_location2 value
        if file_location1 and file_location2: # if the file_location value is not empty
            try: # try to read the data from Excel sheet
                file1, file2 = pd.read_excel(file_location1, dtype={'ID': int,'CustomerID': int, 'Total_Cost': int, 'Level': str}), pd.read_excel(file_location2, dtype={'ID': int, 'CustomerID': int, 'Total_Cost': int, 'Level': str}) # read the data from Excel sheet
                buffer = BytesIO() # create a buffer to store the PDF
                c = canvas.Canvas(buffer, pagesize=letter) # create a canvas to draw the PDF

                y,x1, x2, x3, x4= 700, 100, 200, 300, 400 # create the coordinates of the circles

                combined_files = pd.concat([file1, file2], axis = 1) # combine the two dataframes
                combined_files_selected = combined_files.loc[:, ["ID", "CustomerID", "Total_Cost", "Level"]] # remove the duplicate columns
                combined_files_selected = combined_files_selected.loc[:, ~combined_files_selected.columns.duplicated(keep='first')] # remove the duplicate rows

                # Add column names
                c.drawString(x1, y, "ID")
                c.drawString(x2, y, "Customer ID")
                c.drawString(x3, y, "Total Cost")
                c.drawString(x4, y, "Level")
                y -= 20 # Change the y coordinate

                for _, row in combined_files_selected.iterrows(): # loop through the rows of the data
                    row_id = row['ID'] # get the ID
                    row_customer_id = row['CustomerID'] # get the row_customer_id
                    row_total_cost = row['Total_Cost'] # get the row_total_cost
                    row_level = row['Level'] # get the level
                    # Append data to PDF
                    c.drawString(x1, y, str(row_id))
                    c.drawString(x2, y, str(row_customer_id))
                    c.drawString(x3, y, str(row_total_cost))
                    c.drawString(x4, y, str(row_level))
                    y -= 20

                c.save() # Save the PDF
                try: # try to save the data 
                    name_of_printer = win32print.GetDefaultPrinter() # get the printer name
                    hprinter = win32print.OpenPrinter(name_of_printer) # open the printer
                    try: # try to save the data
                        buffer.seek(0) # reset the buffer
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp: # create a temporary file
                            tmp.write(buffer.read()) # write the buffer to the temporary file
                            tmp.flush() # flush the temporary file
                            os.startfile(tmp.name) # start the temporary file
                        buffer.close() # close the buffer
                        win32print.ClosePrinter(hprinter) # close the printer
                    except Exception as e: # if the printer cannot be opened
                        print("Error:", e) # print the error
                except Exception as e: 
                    print("Error:", e)

            except Exception as e:
                tk.messagebox.showerror("Error", str(e))
    def save_table_data(self):# create a function to save the table data
        # Below opens a filedialog and asks where to save the file in the desktop, there are filetypes that accepts only excel files to be saved
        if not (
            popup_filename := filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")],
            )
        ):
            return
        
        workbook = openpyxl.Workbook() # create a new workbook
        sheet = workbook.active # get the active sheet
        file_location1 = self.file_location.get() # get the file_location value
        file_location2 = self.file_location2.get() # get the file_location2 value
        if file_location1 and file_location2: # if the file_location value is not empty
            try: # try to read the data from Excel sheet
                file1, file2 = pd.read_excel(file_location1, dtype={ 'ID': float, 'CustomerID': int, 'Total_Cost': float}), pd.read_excel(file_location2, dtype={ 'ID': float, 'CustomerID': int, 'Total_Cost': float}) # read the data from Excel sheet
            except Exception as e: # if the file cannot be read
                tk.messagebox.showerror("Error", str(e)) # show the error
                return # return if the file cannot be read
            combined_files = pd.concat([file1, file2], axis=1) # add the combined files to the dataframe
            name_of_column = self.entry_for_column.get() # get the name of the column
            if name_of_column and name_of_column in combined_files.columns: # if the name of the column is valid
                columns_to_save = ["ID", "Total_Cost"] # get the columns to save
                empty_data = [columns_to_save] # create an empty list to store the data
                # Iterate over the rows of the dataframe and append valid data to the output list
                for index, row in combined_files[columns_to_save].iterrows(): # loop through the rows of the dataframe
                    id_value = row["ID"] # get the ID
                    total_cost_value = row["Total_Cost"] # get the row_total_cost
                    if pd.notnull(id_value) and pd.notnull(total_cost_value): # if the ID and Total_Cost are not null
                        empty_data.append([int(id_value), float(total_cost_value)]) # append the data to the output list
               
                # Append data to the sheet
                for row in empty_data:
                    sheet.append(row)

            # save the workbook to the selected filename
            workbook.save(popup_filename)
            messagebox.showinfo("Success", f"Data saved to {popup_filename}")
        else:
            messagebox.showwarning("Warning", "Wrong column name or missing columns")
    def more_rows(self):# create a function to save the table data
        self.btn_more_clicks += 1 # increment the button_more_clicks value
        file_location1 = self.file_location.get() # get the file_location value
        file_location2 = self.file_location2.get() # get the file_location2 value
        if file_location1 and file_location2: # if the file_location value is not empty
            try: # try to read the data from Excel sheet
                file1, file2 = pd.read_excel(file_location1, dtype={'ID': float, 'CustomerID': int, 'Total_Cost': float, 'Level': str}), pd.read_excel(file_location2, dtype={'ID': float, 'CustomerID': int, 'Total_Cost': float, 'Level': str}) # read the data from Excel sheet
            except Exception as e: # if the file cannot be read
                tk.messagebox.showerror("Error", str(e)) # show the error
                return # return if the file cannot be read
            combined_files = pd.concat([file1, file2], ignore_index=True) # add the combined files to the dataframe
            name_of_column = self.entry_for_column.get() # get the name of the column
            if (name_of_column and name_of_column in combined_files.columns and name_of_column == 'Total_Cost'): # if the name of the column is valid
                new_row_data = [
                    ("", i * 1000)
                    for i in range(
                        self.btn_more_clicks * 8 + 1,
                        (self.btn_more_clicks + 1) * 8 + 1,)
                ] # create an empty list to store the data
                for row in new_row_data: # Iterate over the rows of the dataframe and append valid data to the output list
                    self.treeview.insert("", "end", values=row) # append the data to the output list
                self.total_cost_color_updated() # update the total_cost color
    def total_cost_color_updated(self): # update the total cost color
        counts_of_colors = {'red': 0, 'green': 0, 'grey': 0} # create an empty dictionary to store the counts of colors
        index_total_cost = self.treeview["columns"].index("Total_Cost") # get the index of the total_cost column
        tag_for_red, tag_for_green, tag_for_grey = "red_row", "green_row", "grey_row" # create tags for red, green, and grey
        
        self.treeview.tag_configure(tag_for_red, background='', foreground='') # Clear red tags
        self.treeview.tag_configure(tag_for_green, background='', foreground='') # Clear green tags
        self.treeview.tag_configure(tag_for_grey, background='', foreground='') # Clear grey tags
        
        for row in self.treeview.tag_has(tag_for_grey): # Iterate over the rows of the dataframe
            self.treeview.item(row, tags=('untagged',)) # Clear the untagged rows
        for row in self.treeview.get_children(): # Iterate over the rows of the dataframe
            total_cost = self.treeview.set(row, index_total_cost) # get the total_cost value
            if total_cost == "": # if the total_cost value is empty
                continue # if the total_cost value is empty
            total_cost = float(total_cost) # get the total_cost value
            if total_cost >= 5000: # if the total_cost value is greater than 5000
                self.treeview.item(row, tags=(tag_for_red, "total_cost")) # append the total_cost value to the red tag
                counts_of_colors['red'] += 1 # increment the counts of colors
            elif total_cost <= 1000: # if the total_cost value is less than 1000
                self.treeview.item(row, tags=(tag_for_grey, "total_cost")) # append the total_cost value to the grey tag
                counts_of_colors['grey'] += 1 # increment the counts of colors 
            elif 1000 <= total_cost < 5000: # if the total_cost value is greater than
                self.treeview.item(row, tags=(tag_for_green, "total_cost")) # append the total_cost value to the green tag
                counts_of_colors['green'] += 1 # increment the counts of colors

        # show the red tag
        if self.color_data.get() == 'red':
            self.treeview.tag_configure(tag_for_red, background='red', foreground='white')
        # show the green tag
        if self.color_data.get() == 'green':
            self.treeview.tag_configure(tag_for_green, background='green', foreground='white')
        # show the grey tag
        if self.color_data.get() == 'grey':
            self.treeview.tag_configure(tag_for_grey, background='grey', foreground='white')
        
        for row in self.treeview.get_children(): # Iterate over the rows of the dataframe
            total_cost = self.treeview.set(row, index_total_cost) # get the total_cost value
            if total_cost == "": # if the total_cost value is empty
                continue #  if the total_cost value is empty
            total_cost = float(total_cost) # get the total_cost value
            if total_cost >= 5000: # if the total_cost value is greater than 5000
                self.treeview.item(row, tags=(tag_for_red, "total_cost")) # append the total_cost value to the red tag
            elif total_cost <= 1000: # if the total_cost value is less than 1000
                self.treeview.item(row, tags=(tag_for_grey, "total_cost")) # append the total_cost value to the grey tag
            elif 1000 <= total_cost < 5000: # checking if the total_cost value is greater than 5000 or less than 1000 
                self.treeview.item(row, tags=(tag_for_green, "total_cost")) # append the total_cost value to the green tag
    def level_filter_data(self, event=None): # filter the data 
        file_location1 = self.file_location.get() # get the file_location value
        file_location2 = self.file_location2.get() # get the file_location2 value
        selected_column = self.entry_for_column.get() # get the name of the column to focus on LEVEL column 
        if selected_column != "Level": 
            return
        if file_location1 and file_location2: # if the file_location value is not empty
            try: # try to read the data from Excel sheet
                file1, file2 = pd.read_excel(file_location1, dtype={'ID': float, 'Total_Cost': float, 'Level': str, 'CustomerID': int}), pd.read_excel(file_location2, dtype={'ID': float, 'Total_Cost': float, 'Level': str, 'CustomerID': int}) # read the data from Excel sheet
            except Exception as e: # if the file cannot be read
                tk.messagebox.showerror("Error", str(e)) # show the error
                return # return if the file cannot be read
            combined_files = pd.concat([file1, file2], ignore_index=True) # add the combined files to the dataframe
            name_of_column = self.entry_for_column.get() # get the name of the column
            if name_of_column and name_of_column in combined_files.columns: # if the name of the column is valid
                combined_files['Level'] = combined_files['Total_Cost'].apply(lambda x: 1 if x >= 5000 else 0) # filter the data
                if self.level.get() == 1: # if the level is 1
                    combined_files = combined_files[combined_files['Level'] == 1] # filter the data
                elif self.level.get() == 0: # if the level is 0
                    combined_files = combined_files[combined_files['Level'] == 0] # filter the data

                # clear the treeview
                for child in self.treeview.get_children():
                    self.treeview.delete(child)

                # populate the treeview with filtered data
                for i, row in combined_files.iterrows():
                    with contextlib.suppress(ValueError):
                        if self.level.get() == 1 and row['Level'] == 1:
                            values = ['', int(row['Total_Cost']), row['Level']]
                            self.treeview.insert("", "end", values=values)
                        elif self.level.get() == 0 and row['Level'] == 0:
                            values = ['', int(row['Total_Cost']), row['Level']]
                            self.treeview.insert("", "end", values=values)
    def change_popup_window(self): # create a popup window
        top = Toplevel() # create a window
        top.geometry("250x100") # set the size of the window
        self.new_value = Entry(top, width=25, relief=tk.SUNKEN, borderwidth=5, font=("Arial Bold", 10), fg="#02f2b6", bg="#808080") # create a new entry
        self.new_value.pack(pady = 10, side = TOP) # pack the new entry
        self.change_button = tk.Button(top, text="Change", command=self.data_modify, width=8, height=1, borderwidth=7, bg="red", fg="white", font=("Arial Bold", 11)) # create a new change button 
        self.change_button.place(x = 76, y= 45) # place the change button 
        top.mainloop() # show the window
    def data_modify(self): # modify the data
        file_location1, file_location2= self.file_location.get(), self.file_location2.get() # get the file_location value
        new_value, new_value2, selected_item  = self.new_value.get(), self.new_value.get(), self.treeview.focus() # get the new_value value and the new_value2 value and the selected_item value
        workbook_loaded, workbook_loaded2 = openpyxl.load_workbook(file_location1), openpyxl.load_workbook(file_location2)# load the workbook
        ws, ws2, row_number= workbook_loaded.active, workbook_loaded2.active, self.treeview.index(selected_item) # get the worksheets and the row number
        
        if file_location1 and file_location2: # if the file_location value is not empty
            try: # try to read the data from Excel sheet
                file1, file2 = pd.read_excel(file_location1, dtype={'ID': float, 'Total_Cost': float, 'Level': str, 'CustomerID': int}), pd.read_excel(file_location2, dtype={'ID': float, 'Total_Cost': float, 'Level': str, 'CustomerID': int}) # read the data from Excel sheet
            except Exception as e: # if the file cannot be read
                tk.messagebox.showerror("Error", str(e)) # show the error
                return # return if the file cannot be read
            combined_files = pd.concat([file1, file2], ignore_index=True) # add the combined files to the dataframe
            name_of_column = self.entry_for_column.get() # get the name of the column
            if name_of_column and name_of_column in combined_files.columns: # if the name of the column is valid
                if selected_item: # if the selected_item value is not empty
                    if name_of_column == 'ID': # if the name of the column is ID
                        combined_files.at[selected_item, 'ID'] = new_value # update the ID value
                        self.treeview.set(selected_item, column="ID", value=new_value2) # update the ID value
                        column_number = combined_files.columns.get_loc('ID')-4 # get the column number
                        combined_files.iloc[row_number, column_number] = new_value # update the ID value
                        ws2.cell(row=row_number+2, column=column_number, value=new_value2) # update the ID value
                        workbook_loaded2.save(file_location2) # save the updated workbook
                    elif name_of_column == 'Total_Cost': # if the name of the column is Total_Cost
                        combined_files.at[selected_item, 'Total_Cost'] = new_value # update the Total_Cost value
                        self.treeview.set(selected_item, column="Total_Cost", value=new_value) # update the Total_Cost value
                        self.display_info.delete(0, tk.END) # clear the display frame
                        self.display_info.insert(tk.END, f"Total Cost: {new_value}") # update the display frame
                        column_number = combined_files.columns.get_loc('Total_Cost')+1 # get the column number
                        combined_files.iloc[row_number, column_number] = new_value # update the Total_Cost value
                        ws.cell(row=row_number+2, column=column_number, value=new_value) # update the Total_Cost value
                        workbook_loaded.save(file_location1) # save the updated workbook
                else: # if the selected_item value is empty
                    messagebox.showerror("Error", "Select an item from the table.") # show the error
                    return None # return if the selected_item value is empty
            else: # if the name of the column is not valid
                messagebox.showerror("Error", "Select a valid column.") # show the error
                return None # return if the name of the column is not valid 
    def search_data_button(self): #search the data
        file_location1, file_location2 = self.file_location.get(), self.file_location2.get() #  get the file location
        if file_location1 and file_location2: # if the file location value is not empty
            try: # try to read the data from Excel sheet
                file1, file2= pd.read_excel(file_location1, dtype={'ID': float, 'CustomerID': int, 'Total_Cost': float, 'Level': str}), pd.read_excel(file_location2, dtype={'ID': float, 'CustomerID': int, 'Total_Cost': float, 'Level': str}) # read the data from Excel sheet
            except Exception as e: # if the file cannot be read
                tk.messagebox.showerror("Error", str(e)) # show the error
                return # return if the file cannot be read
            combined_files, column_name = pd.concat([file1, file2], ignore_index=True), self.entry_for_column.get() # add the combined files to the dataframe
            
            if column_name and column_name in combined_files.columns: # if the name of the column is valid
                if column_name == 'Level': # if the name of the column is Level
                    filtered_data = combined_files[['Level']].dropna(how='all') # drop the duplicate rows
                else: # if the name of the column is not Level
                    filtered_data = combined_files[[column_name]].drop_duplicates().dropna(subset=[column_name]) # drop the duplicate rows
                for item in self.treeview.get_children(): # for each item in the treeview
                    self.treeview.delete(item) # delete the item
                for i, row in filtered_data.iterrows(): # for each row in the filtered data
                    list_of_values = [] # create a list of list_of_values
                    for col in self.treeview['columns']: # for each column in the treeview
                        if col in row.index: # if the column is in the row
                            if col == 'Level': # if the column is Level
                                value = row[col] # get the value
                                if pd.isnull(value) or value == '': # if the value is empty
                                    list_of_values.append('') # add an empty string
                                else: # if the value is not empty
                                    list_of_values.append(str(value)) # add the value
                            elif col in ['Total_Cost', 'ID']: # if the column is Total_Cost or ID 
                                value = row[col] # get the value
                                if pd.isnull(value) or value == '': # if the value is empty
                                    list_of_values.append('0') # add an empty string
                                else:   # if the value is not empty
                                    list_of_values.append(int(value)) # add the value
                            else: # if the column is not Level or Total_Cost
                                list_of_values.append(row[col]) # add the value   
                        else: # if the column is not in the row
                            list_of_values.append('') # add an empty string
                    self.treeview.insert("", "end", values=list_of_values) # add the values to the treeview

                index_total_cost = self.treeview["columns"].index("Total_Cost") # get the index of Total_Cost
                for row in self.treeview.get_children():  # for each item in the treeview
                    total_cost = self.treeview.set(row, index_total_cost) # get the value of Total_Cost
                    if total_cost.strip(): # if the value of Total_Cost is not empty
                        if int(total_cost) >= 5000: # if the value of Total_Cost is greater than 5000
                            if self.color_data.get() == 'red': # if the color is red
                                self.treeview.item(row, tags=("red_row",)) # set the tag to red
                        elif int(total_cost) <= 1000: # if the value of Total_Cost is less than 1000
                            if self.color_data.get() == 'grey': # if the color is grey
                                self.treeview.item(row, tags=("grey_row",)) # set the tag to grey
                        elif 1000 <= int(total_cost) <= 5000: # if the value of Total_Cost is between 1000 and 5000
                            if self.color_data.get() == 'green': # if the color is green
                                self.treeview.item(row, tags=("green_row",)) # set the tag to green
            else: # if the name of the column is not valid
                tk.messagebox.showwarning("Warning", "Wrong Column Name") # show the warning
        else:
            tk.messagebox.showwarning("Warning", "Select both files") # show the warning 
    def file_browse(self): # open a file dialog to get the file path
        file_location = filedialog.askopenfilename() # get the file path
        self.file_location.delete(0, tk.END) # clear the path entry
        self.file_location.insert(0, file_location) # insert the file path
    def file_browse2(self): # open a file dialog to get the file path
        for tag in ['red_row', 'green_row', 'grey_row']:
            self.treeview.tag_configure(tag, background='', foreground='') # clear the tags

        file_location = filedialog.askopenfilename() # get the file path

        self.df = pd.read_excel(file_location, engine='openpyxl') # read the data from Excel sheet
        self.file_location2.delete(0, 'end') # clear the path entry
        self.file_location2.insert('end', file_location) # insert the file path
        self.data_loaded = True # set the data loaded to True
        self.treeview.delete(*self.treeview.get_children()) # clear the treeview
    def show_data_message(self): # show the data in the treeview
        if file_location := self.file_location.get(): # if the file location value is not empty
            try: # try to read the data from Excel sheet
                file1 = pd.read_excel(file_location, dtype={'ID': float, 'CustomerID': int, 'Total_Cost': int, 'Level': str}) # read the data from Excel sheet
                file1['ID'] = file1['ID'].astype(int) # convert the 'ID' column to integer
            except Exception as e: # if the file cannot be read
                tk.messagebox.showerror("Error", str(e)) # show the error
                return # return if the file cannot be read
            name_of_column = self.entry_for_column.get() # get the column name entered by the user
            if name_of_column and name_of_column in file1.columns: # if the name of the column is valid
                filtered_data = file1[[name_of_column]] # filter the data based on the entered column name and remove duplicates
                self.display_data_frame(filtered_data) # display the filtered data
            else: # if the name of the column is not valid
                tk.messagebox.showwarning("Warning", "Wrong column name") # show the warning
    def show_data_message2(self): # show the data in the treeview for file2 
        if not (file_location := self.file_location2.get()): # if the file location value is empty
            return # return if the file location value is empty
        try: # try to read the data from Excel sheet
            file2 = pd.read_excel(file_location, dtype={'ID': int, 'CustomerID': int, 'Total_Cost': int, 'Level': str}) # read the data from Excel sheet
            file2['ID'] = file2['ID'].astype(int) # convert the 'ID' column to integer
        except Exception as e: # if the file cannot be read
            tk.messagebox.showerror("Error", str(e)) # show the error
            return # return if the file cannot be read
        name_of_column = self.entry_for_column.get() # get the column name entered by the user
        if name_of_column and name_of_column in file2.columns: # if the name of the column is valid
            filtered_data = file2[[name_of_column]] # filter the data based on the entered column name and remove duplicates
            for row in self.treeview.get_children(): # for each item in the treeview
                self.treeview.delete(row) # delete the item
            for i, row in filtered_data.iterrows(): # for each row in the filtered data
                values = [] # create a list of values
                for col in self.treeview['columns']: # for each column in the treeview
                    if col in row.index: # if the column is in the row
                        values.append(row[col]) # add the value
                    else: # if the column is not in the row
                        values.append('') # add an empty string
                self.treeview.insert('', 'end', values=values) # add the values to the treeview
        else: # if the name of the column is not valid
            tk.messagebox.showwarning("Warning", "Wrong column name") # show the warning 
    def display_data_frame(self, data): # display the data in the treeview
        for widget in self.display_info.winfo_children(): # for each widget in the display_
            widget.destroy() # destroy the widget
        data_columns, treeview = list(data.columns), tk.ttk.Treeview(self.display_info, columns=data_columns, show="headings") # create the treeview
        for col in data_columns: # for each column in the data
            treeview.heading(col, text=col) # add the column heading to the treeview
        for i, row in data.iterrows(): # for each row in the data
            values = [row[col] for col in data_columns] # create a list of values
            treeview.insert("", "end", text=i+1, values=values) # add the values to the treeview
        index_total_cost = data_columns.index('Total_Cost') # get the index of Total_Cost
        for i, row in data.iterrows(): # for each row in the data
            total_cost_value = row['Total_Cost'] # get the value of Total_Cost
            treeview.set(treeview.get_children()[i], data_columns[index_total_cost], total_cost_value) # set the value of Total_Cost
        treeview.pack() # pack the treeview


if __name__ == "__main__": # if the file is executed directly from the command line then execute the main function
    root = tk.Tk() # create the root window
    root.geometry("1100x700") # set the size of the window
    app = database_management(root) # create the app
    app.pack(expand=True, fill="both") # pack the app
    root.mainloop() # start the main loop
