# =============================================================================
# File Name: Mapside Mapper.py
# File Version: 0.1.8
# Python Version: 3.11
# =============================================================================
# Created By: Clayton Barnett
# Date Created: 2/9/2021
# Date Last Modified: 7/26/2023
# =============================================================================

import os
import numpy as np
import pandas as pd
import tkinter as tk
import tkinter.ttk as ttk
import simplekml
from tkinter import * 
from ttkthemes import ThemedStyle
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
from geographiclib.geodesic import Geodesic
from geopy.distance import geodesic
from pyproj import Geod

# Get the desktop, documents, and coordinates file paths
desktop_path = os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop")
documents_path = os.path.join(os.path.expanduser("~"), "OneDrive", "Documents")
coordinates_path = os.path.join(os.path.expanduser("~"), "OneDrive", "Google Earth", "Google Earth - Job Alignment Coordinates")

class ToolTip(object):
    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0

    def showtip(self, text):
        "Display text in tooltip window"
        self.text = text
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 57
        y = y + cy + self.widget.winfo_rooty() +27
        self.tipwindow = tw = Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = Label(tw, text=self.text, justify=LEFT,background="#ffffff", relief=SOLID, borderwidth=1,
        font=("Montserrat", "8", "normal"))
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

def CreateToolTip(widget, text):
    toolTip = ToolTip(widget)
    def enter(event):
        toolTip.showtip(text)
    def leave(event):
        toolTip.hidetip()
    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)

# Working Code =============================================================================
def import_function():
    global data
    global first_station
    global first_station_int
    global last_station
    global last_station_int
    Tk().withdraw()

    # Selecting excel file to import as dataframe
    infile = askopenfilename(initialdir = coordinates_path, title="Select Coordinates File", filetypes=[
        ('Excel Files (*.csv *.xls *.xlsx *.xlsm)', '*.csv *.xlsm *.xlsx *.xls'),
        ('All Files (*.*)', '*'), ])
    
    # Getting file extension
    ext = os.path.splitext(infile)[1]
    # If *.csv file, use read_csv
    if ext == '.csv':
        df = pd.read_csv(infile)
    # If excel file, use read_excel
    elif ext == '.xls' or '.xlsx' or '.xlsm':
        df = pd.read_excel(infile)
    # If none of these file extensions, raise RunTimeError
    else:
        raise RuntimeError('File extension not recognized')
    # Define the ellipsoid (How to Calculate)
    geod = Geod("+ellps=WGS84")

    # Creating variables for Start (x) and End (y) Coordinates
    x = 0
    y = 1

# Getting Number of Bends =============================================================================
    df_len = len(df)
    distancelist=[]
    dlist=[]
    int_dlist=[]

# Getting Total Line Distance (What the length should be) =============================================================================
    for i in range(len(df)-1):

        # Creating start and end variables to get the coordinates between them
        start = df.Latitude[x], df.Longitude[x] # After each iteration x moves down to the next row (x += 1)
        end = df.Latitude[y], df.Longitude[y] # After each iteration y moves down to the next row (y += 1)

        # Measuring distance between start and end variables, round decimals and return as integer
        distance_float = round((geodesic(start, end).ft),2)
        distance = int(distance_float)
        int_dlist.append(distance)
        dlist.append(distance_float)
        distancelist.append(distance)

        # Adding an increment of 1 to variables 'x' and 'y' to move to the next row
        x += 1
        y += 1

    actual_length = int(sum(dlist))
    # Total_Line_Length = int(sum(dlist))
    distance_sum = round(sum(distancelist))
    last_station_num = df['Name'].iloc[-1]
    last_station_num = int(last_station_num.replace('+', ''))


# Getting base program's incorrect result =============================================================================

    # Creating variables for Start (x) and End (y) Coordinates
    x = 0
    y = 1

    # Creating nested list 'nlist' here and append each element to it during iteration
    nlist=[]
    wrong_distance_list = []

    # Creating a for loop to iterate through the length of the dataframe
    for i in range(len(df)-1):
        # Setting start_coord to first index 0 and then selecting next row after each iteration
        start_coord = df['Longitude'][x], df['Latitude'][x]

        # Creating start and end variables to get the coordinates between them
        start = df.Latitude[x], df.Longitude[x] # After each iteration x moves down to the next row (x += 1)
        end = df.Latitude[y], df.Longitude[y] # After each iteration y moves down to the next row (y += 1)

        # Measuring distance between start and end variables, round decimals and return as integer
        distance = round((geodesic(start, end).ft), 2)

        # Getting number of points (npts) between lat1/lon1 and lat2/lon2 (start and end) incrementally moving down the rows of the dataframe
        # npts is the integer number of the distance variable and saving to list 'a'
        a = geod.npts(
            lat1=df['Latitude'][x],
            lon1=df['Longitude'][x],
            lat2=df['Latitude'][y],
            lon2=df['Longitude'][y],
            npts=distance) # If distance is 100 feet, there will be 100 points or 1 point every foot
        
        # Inserting start_coord to the first element of list 'a'
        a.insert(0, start_coord)

        # Adding list 'a' to nested list 'nlist'
        nlist.append(a)


        # Adding an increment of 1 to variables 'x' and 'y' to move to the next row
        x += 1
        y += 1

    # Flattening the nested list 'nlist' to a single list instead of a list of lists
    # Without this, the nlist will only output the last list added from list 'a'
    flat_list = [item for sublist in nlist for item in sublist]

    # Adding the last coordinate of the line to the end of the new list
    end_of_line_coord = df['Longitude'].iloc[-1], df['Latitude'].iloc[-1]
    flat_list.append(end_of_line_coord)

    # Adding flattened list to a new dataframe
    data = pd.DataFrame(flat_list)
    rename_index = str(df['Name'][0])
    rename_index = int(rename_index.replace('+', ''))
    data.index += rename_index

    station_difference = actual_length - distance_sum
    minus_number = float(station_difference / df_len)

    data.to_csv(f'{documents_path}\\Mapside Mapping Coordinates Incorrect.csv', sep=',', index=False)
    data = pd.DataFrame()

# --------Getting base program's correct final result----------------------------------------------------------

    # Creating variables for Start (x) and End (y) Coordinates
    x = 0
    y = 1

    # Creating nested list 'nlist' here and append each element to it during iteration
    nlist=[]
    correct_distance_list = []

    # Creating a for loop to iterate through the length of the dataframe
    for i in range(len(df)-1):
        # Setting start_coord to first index 0 and then selecting next row after each iteration
        start_coord = df['Longitude'][x], df['Latitude'][x]

        # Creating start and end variables to get the coordinates between them
        start = df.Latitude[x], df.Longitude[x] # After each iteration x moves down to the next row (x += 1)
        end = df.Latitude[y], df.Longitude[y] # After each iteration y moves down to the next row (y += 1)

        # Measuring distance between start and end variables, round decimals and return as integer
        distance = (round((geodesic(start, end).ft), 2)-minus_number)
        
        # Getting number of points (npts) between lat1/lon1 and lat2/lon2 (start and end) incrementally moving down the rows of the dataframe
        # npts is the integer number of the distance variable and saving to list 'a'
        a = geod.npts(
            lat1=df['Latitude'][x],
            lon1=df['Longitude'][x],
            lat2=df['Latitude'][y],
            lon2=df['Longitude'][y],
            npts=distance) # If distance is 100 feet, there will be 100 points or 1 point every foot
        
        # Inserting start_coord to the first element of list 'a'
        a.insert(0, start_coord)

        # Adding list 'a' to nested list 'nlist'
        nlist.append(a)
        correct_distance_list.append(distance)

        # Adding an increment of 1 to variables 'x' and 'y' to move to the next row
        x += 1
        y += 1

    # Flattening the nested list 'nlist' to a single list instead of a list of lists
    # Without this, the nlist will only output the last list added from list 'a'
    flat_list = [item for sublist in nlist for item in sublist]

    # Adding the last coordinate of the line to the end of the new list
    end_of_line_coord = df['Longitude'].iloc[-1], df['Latitude'].iloc[-1]
    flat_list.append(end_of_line_coord)

    # Adding flattened list to a new dataframe
    data = pd.DataFrame(flat_list)
    rename_index = str(df['Name'][0])
    rename_index = int(rename_index.replace('+', ''))
    data.index += rename_index

    # Dropping any duplicates from new dataframe
    data.drop_duplicates(keep=False, inplace=True)

    # Naming first two columns as Longitude and Latitude
    data = data.rename(columns={data.columns[0]:'Longitude', data.columns[1]: 'Latitude'})

    # Creating variable for desired column layout
    columns_titles = ["Latitude","Longitude"]

    # Reindexing columns to the desired column layout
    data=data.reindex(columns=columns_titles)

    # Saving the dataframe as a .csv file and separarating values by comma if some coordinates weren't separated
    data.to_csv(f'{documents_path}\\Mapside Mapping Coordinates Correct.csv', sep=',', index=True)

    first_station = str(df['Name'].iloc[0])
    first_station_int = int(first_station.replace('+',''))
    last_station = df['Name'].iloc[-1]
    last_station_int = int(last_station.replace('+',''))

    # Creating a label in the lower left corner to show supported file types
    upper_left = tk.Label(root, bg='#D8D8D8', font=('calibre',10), text=(f'Valid Station Numbers: {first_station} to {last_station}'))
    upper_left.place(relx=0.5, rely=0, anchor='n')

def help_message():
    messagebox.showinfo(title='Help Menu', message='To Get Started Click On:\nFile > Import > Select A Coordinates Excel File')

# Creating Tk window
root = tk.Tk()
root.configure(background='#D8D8D8')

# Setting Theme
style = ThemedStyle(root)
style.set_theme("clam")
root.iconbitmap(f"{documents_path}\\Mapside Program\\assets\\Mapside.ico")

# Setting the windows size
root.title("Mapside Mapper")  # should always be inside mainloop
root.geometry("350x200") # set window size
root.minsize(350, 200)
root.maxsize(350, 200)

# Creating Menu
menu = Menu(root)

# Adding File Menu and commands 
file_menu = Menu(menu, tearoff=0)
menu.add_cascade(label='File', menu=file_menu)
file_menu.add_command(label='Import...', command=import_function)
file_menu.add_separator()
file_menu.add_command(label='Exit', command=root.destroy)

# Adding Help Menu 
help_menu = Menu(menu, tearoff=0)
menu.add_cascade(label='Help', menu=help_menu)
help_menu.add_command(label='Help', command=help_message)
help_menu.add_command(label='About', command=None)
help_menu.add_separator()
help_menu.add_command(label='Exit', command=root.destroy)

# Display Menu
root.config(menu=menu)

# Declaring string variable for storing "From Station:" and "To Station:" inputs
from_input_var=tk.StringVar()
to_input_var=tk.StringVar()

def submit1(): # Click submit button
    f=[]
    from_sta = from_input_var.get()
    to_sta = to_input_var.get()

    from_num = int(from_sta)
    from_num_list = str(from_num)
    from_num_list = list(from_num_list)
    if from_num >= 0 and from_num <= 9:
        from_num_list.insert(0, '00+0')
        from_num_list = ''.join(from_num_list)
    elif from_num >= 10 and from_num <= 99:
        from_num_list.insert(0, '00+')
        from_num_list = ''.join(from_num_list)
    elif from_num >= 100 and from_num <= 999:
        from_num_list.insert(0, '0')
        from_num_list.insert(2, '+')
        from_num_list = ''.join(from_num_list)
    elif from_num >= 1000 and from_num <= 9999:
        from_num_list.insert(2, '+')
        from_num_list = ''.join(from_num_list)
    elif from_num >= 10000 and from_num <= 99999:
        from_num_list.insert(3, '+')
        from_num_list = ''.join(from_num_list)
    elif from_num >= 100000 and from_num <= 999999:
        from_num_list.insert(4, '+')
        from_num_list = ''.join(from_num_list)
    elif from_num >= 1000000 and from_num <= 9999999:
        from_num_list.insert(5, '+')
        from_num_list = ''.join(from_num_list)
        

    to_num = int(to_sta)
    to_num_list = str(to_num)
    to_num_list = list(to_num_list)
    if to_num >= 0 and to_num <= 9:
        to_num_list.insert(0, '00+0')
        to_num_list = ''.join(to_num_list)
    elif to_num >= 10 and to_num <= 99:
        to_num_list.insert(0, '00+')
        to_num_list = ''.join(to_num_list)
    elif to_num >= 100 and to_num <= 999:
        to_num_list.insert(0, '0')
        to_num_list.insert(2, '+')
        to_num_list = ''.join(to_num_list)
    elif to_num >= 1000 and to_num <= 9999:
        to_num_list.insert(2, '+')
        to_num_list = ''.join(to_num_list)
    elif to_num >= 10000 and to_num <= 99999:
        to_num_list.insert(3, '+')
        to_num_list = ''.join(to_num_list)
    elif to_num >= 100000 and to_num <= 999999:
        to_num_list.insert(4, '+')
        to_num_list = ''.join(to_num_list)
    elif to_num >= 1000000 and to_num <= 9999999:
        to_num_list.insert(5, '+')
        to_num_list = ''.join(to_num_list)

    from_sta = int(from_sta)
    to_sta = int(to_sta)
    
    for lat, lon in zip(data["Latitude"].loc[from_sta:to_sta], data["Longitude"].loc[from_sta:to_sta]):
        coord = (lon, lat)
        f.append(coord)
    test_coord = np.array(f)

    # Checking for incorrect input in 'from station' entry box
    # If from station is less than the lowest station number, show an error
    if from_sta < first_station_int:
        tk.messagebox.showerror(title='Distance Error', message=f'From Station Number cannot be less than {first_station}.')
        raise RuntimeError('FromStation_NotValid')
    # If from station is greater than the highest station number, show an error
    elif from_sta > last_station_int:
        tk.messagebox.showerror(title='Distance Error', message=f'From Station Number cannot be higher than {last_station}.')
        raise RuntimeError('FromStation_NotValid')

    # Checking for incorrect input in 'to station' entry box
    # If from station is less than the lowest station number, show an error
    if to_sta < first_station_int:
        tk.messagebox.showerror(title='Distance Error', message=f'To Station Number cannot be less than {first_station}.')
        raise RuntimeError('ToStation_NotValid')
    # If from station is greater than the highest station number, show an error
    elif to_sta > last_station_int:
        tk.messagebox.showerror(title='Distance Error', message=f'To Station Number cannot be higher than {last_station}.')
        raise RuntimeError('ToStation_NotValid')
    # Checking if from station is higher than to station
    # If 'from station' is greater than 'to station', show an error
    if from_sta > to_sta:
        tk.messagebox.showerror(title='Distance Error', message='From Sta # must be smaller than To Sta #')
        raise RuntimeError('FromStation_GreaterThan_ToStation')
    # Checking if total line length is greater than 65,521.
    # 65,521 is the maximum amount of points a kmz line can have before it does not display in Google Earth
    # If total line length is greater than 65,521, show an error
    if from_and_to_diff >= 65521:
        tk.messagebox.showerror(title='Distance Error', message='Length must not exceed 65,521 feet\'')
        raise RuntimeError('Length_To_Long')

    from_input_var.set("")
    to_input_var.set("")

    kml = simplekml.Kml(open=1)

    ls = kml.newlinestring(name='New LineString', description=f'{from_num_list} to {to_num_list}')
    # Creating an array for coordinate points
    ls.coords = np.array(f)
    # New line setting to clamp to the ground
    ls.altitudemode = simplekml.AltitudeMode.clamptoground
    # Connects the new line to the ground
    ls.extrude = 1
    # Set new line width to 2.
    ls.style.linestyle.width = 6
    # Set new line color to yellow.
    ls.style.linestyle.color = 'ffff5500'
    line_kmz_path = f'{desktop_path}\\{from_num_list} To {to_num_list}.kmz'
    # Save as a *.kmz file to the selected directory
    kml.save(line_kmz_path)
    print(line_kmz_path)
    # Open with Google Earth Pro
    os.startfile(line_kmz_path)

def submit(event): # Pressing enter key to click on submit button

    f=[]
    from_sta = from_input_var.get()
    to_sta = to_input_var.get()

    from_num = int(from_sta)
    from_num_list = str(from_num)
    from_num_list = list(from_num_list)
    if from_num >= 0 and from_num <= 9:
        from_num_list.insert(0, '00+0')
        from_num_list = ''.join(from_num_list)
    elif from_num >= 10 and from_num <= 99:
        from_num_list.insert(0, '00+')
        from_num_list = ''.join(from_num_list)
    elif from_num >= 100 and from_num <= 999:
        from_num_list.insert(0, '0')
        from_num_list.insert(2, '+')
        from_num_list = ''.join(from_num_list)
    elif from_num >= 1000 and from_num <= 9999:
        from_num_list.insert(2, '+')
        from_num_list = ''.join(from_num_list)
    elif from_num >= 10000 and from_num <= 99999:
        from_num_list.insert(3, '+')
        from_num_list = ''.join(from_num_list)
    elif from_num >= 100000 and from_num <= 999999:
        from_num_list.insert(4, '+')
        from_num_list = ''.join(from_num_list)
    elif from_num >= 1000000 and from_num <= 9999999:
        from_num_list.insert(5, '+')
        from_num_list = ''.join(from_num_list)
        

    to_num = int(to_sta)
    to_num_list = str(to_num)
    to_num_list = list(to_num_list)
    if to_num >= 0 and to_num <= 9:
        to_num_list.insert(0, '00+0')
        to_num_list = ''.join(to_num_list)
    elif to_num >= 10 and to_num <= 99:
        to_num_list.insert(0, '00+')
        to_num_list = ''.join(to_num_list)
    elif to_num >= 100 and to_num <= 999:
        to_num_list.insert(0, '0')
        to_num_list.insert(2, '+')
        to_num_list = ''.join(to_num_list)
    elif to_num >= 1000 and to_num <= 9999:
        to_num_list.insert(2, '+')
        to_num_list = ''.join(to_num_list)
    elif to_num >= 10000 and to_num <= 99999:
        to_num_list.insert(3, '+')
        to_num_list = ''.join(to_num_list)
    elif to_num >= 100000 and to_num <= 999999:
        to_num_list.insert(4, '+')
        to_num_list = ''.join(to_num_list)
    elif to_num >= 1000000 and to_num <= 9999999:
        to_num_list.insert(5, '+')
        to_num_list = ''.join(to_num_list)

    from_sta = int(from_sta)
    to_sta = int(to_sta)

    for lat, lon in zip(data["Latitude"].loc[from_sta:to_sta], data["Longitude"].loc[from_sta:to_sta]):
        coord = (lon, lat)
        f.append(coord)
    test_coord = np.array(f)
    
    from_and_to_diff = int(to_sta) - int(from_sta)

    # Checking for incorrect input in 'from station' entry box
    # If from station is less than the lowest station number, show an error
    if from_sta < first_station_int:
        tk.messagebox.showerror(title='Distance Error', message=f'From Station Number cannot be less than {first_station}.')
        raise RuntimeError('FromStation_NotValid')
    # If from station is greater than the highest station number, show an error
    elif from_sta > last_station_int:
        tk.messagebox.showerror(title='Distance Error', message=f'From Station Number cannot be higher than {last_station}.')
        raise RuntimeError('FromStation_NotValid')

    # Checking for incorrect input in 'to station' entry box
    # If from station is less than the lowest station number, show an error
    if to_sta < first_station_int:
        tk.messagebox.showerror(title='Distance Error', message=f'To Station Number cannot be less than {first_station}.')
        raise RuntimeError('ToStation_NotValid')
    # If from station is greater than the highest station number, show an error
    elif to_sta > last_station_int:
        tk.messagebox.showerror(title='Distance Error', message=f'To Station Number cannot be higher than {last_station}.')
        raise RuntimeError('ToStation_NotValid')
    
    # Checking if from station is higher than to station
    # If 'from station' is greater than 'to station', show an error
    if from_sta > to_sta:
        tk.messagebox.showerror(title='Distance Error', message='From Sta # must be smaller than To Sta #')
        raise RuntimeError('FromStation_GreaterThan_ToStation')
    
    # Checking if total line length is greater than 65,521.
    # 65,521 is the maximum amount of points a kmz line can have before it does not display in Google Earth
    # If total line length is greater than 65,521, show an error
    if from_and_to_diff >= 65521:
        tk.messagebox.showerror(title='Distance Error', message='Length must not exceed 65,521 feet\'')
        raise RuntimeError('Length_To_Long')

    from_input_var.set("")
    to_input_var.set("")

    kml = simplekml.Kml(open=1)
    ls = kml.newlinestring(name='New LineString', description=f'{from_num_list} to {to_num_list}')
    # Creating an array for coordinate points
    ls.coords = np.array(f)
    # New line setting to clamp to the ground
    ls.altitudemode = simplekml.AltitudeMode.clamptoground
    # Connects the new line to the ground
    ls.extrude = 1
    # Set new line width to 2.
    ls.style.linestyle.width = 6
    # Set new line color to yellow.
    ls.style.linestyle.color = 'ffff5500'
    line_kmz_path = f'{desktop_path}\\{from_num_list} To {to_num_list}.kmz'
    # Save as a *.kmz file to the selected directory
    kml.save(line_kmz_path)
    # Open with Google Earth Pro
    os.startfile(line_kmz_path)

root.bind('<Return>', submit)

# Creating a label for "From Station:" using widget Label
from_label = tk.Label(root, text = 'From Station:', font=('Montserrat',10, 'bold'), bg='#D8D8D8')
img = PhotoImage(file=f"{documents_path}\\Mapside Program\\assets\\info.png")
infoLabel1 = ttk.Label(root, image=img)
infoLabel1.place(relx=0.75, rely=0.20, anchor='center')
CreateToolTip(infoLabel1, text = ' Disabled Characters: +, =, ?, / ')

# Creating an entry for input "From Station:" using widget Entry
from_entry = ttk.Entry(root,textvariable = from_input_var, font=('Montserrat',10,'normal'))
from_entry.bind("+", lambda e: "break") # Disable characters from keyboard
from_entry.bind("=", lambda e: "break") # Disable characters from keyboard
from_entry.bind("?", lambda e: "break") # Disable characters from keyboard
from_entry.bind("/", lambda e: "break") # Disable characters from keyboard

# Creating a label for "To Station:"
to_label = ttk.Label(root, text = 'To Station:', font = ('Montserrat',10,'bold'))
infoLabel1 = ttk.Label(root, image=img)
infoLabel1.place(relx=0.75, rely=0.4, anchor='center')
CreateToolTip(infoLabel1, text = ' Disabled Characters: +, =, ?, / ')

# Creating an entry for "To Station:"
to_entry=ttk.Entry(root, textvariable = to_input_var, font = ('Montserrat',10,'normal'))
to_entry.bind("+", lambda e: "break") # Disable characters from keyboard
to_entry.bind("=", lambda e: "break") # Disable characters from keyboard
to_entry.bind("?", lambda e: "break") # Disable characters from keyboard
to_entry.bind("/", lambda e: "break") # Disable characters from keyboard

# Creating a button using the widget Button that will call the submit function 
sub_btn=ttk.Button(root,text='Submit', command=submit1)
import_btn=ttk.Button(root,text = 'Import Coordinates', command = import_function)

# Placing the label and entry in the required position using grid method
from_label.place(relx=0.15, rely=0.20, anchor='center')
from_entry.place(relx=0.50, rely=0.20, anchor='center')
to_label.place(relx=0.15, rely=0.4, anchor='center')
to_entry.place(relx=0.50, rely=0.4, anchor='center')
sub_btn.place(relx=0.5, rely=0.6, anchor='center')
  
# Performing an infinite loop for the window to display
root.mainloop()