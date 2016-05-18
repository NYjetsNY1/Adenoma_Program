# Program Written By Benjamin Sklar
# Adenoma Program
# Version 1.0, very little error checking in this version
# Very quickly created, style not the best. Will attempt to make the code better when there is more time
# to work on this program. Functionally, program works flawlessly.


# necessary imports
import pandas as pd
from tkinter import *
from tkinter import filedialog
from xlrd import open_workbook, XLRDError
from win32com.client import Dispatch as comDispatch


def test_book(filename):
    # try with exception block
    try:
        open_workbook(filename)
    except XLRDError:
        return False
    else:
        return True


def macro_function():
    global macro_name
    macro_name = filedialog.askopenfilename()
    if macro_name.endswith('.XLSB'):

        # Other buttons states
        time_track["state"] = NORMAL
        adenoma_path["state"] = NORMAL
        patient_total["state"] = NORMAL


        # Other button  configurations
        time_track.configure(bg="yellow")
        adenoma_path.configure(bg="yellow")
        patient_total.configure(bg="yellow")

        # Macro button
        up_mac["text"] = "Thank You for Uploading the Macro."
        up_mac.configure(bg="green", fg="white")
        up_mac["state"] = DISABLED


def time_tracking_function():
    time_tracking_file = filedialog.askopenfilename()
    if (test_book(time_tracking_file) == TRUE):
        time_track["text"] = "Thank You for Uploading the Time Tracking Results."
        time_track["state"] = DISABLED
        xl = comDispatch('Excel.Application')
        xl.Workbooks.Open(time_tracking_file, ReadOnly=0)
        xl.Run("'" + macro_name + "'" + "!Time_Tracking")
        xl.Workbooks(1).Close(SaveChanges=1)
        xl.Quit()
        del xl
        time_track.configure(bg = "green")

def adenoma_function():
    adenoma_filename = filedialog.askopenfilename()
    if (test_book(adenoma_filename) == TRUE):
        adenoma_path["text"] = "Thank You for Uploading the Adenoma File from the Path Lab."
        adenoma_path["state"] = DISABLED
        xl = comDispatch('Excel.Application')
        xl.Workbooks.Open(adenoma_filename, ReadOnly=0)
        xl.Run("'" + macro_name + "'" + "!Adenoma_Duplicates")
        xl.Workbooks(1).Close(SaveChanges=1)
        xl.Quit()
        del xl
        adenoma_path.configure(bg="green")


def patient_total_function():
    patient_total_filename = filedialog.askopenfilename()
    if patient_total_filename.endswith('.csv'):
        patient_total["text"] = "Thank You for Uploading the Total Patient File."
        patient_total["state"] = DISABLED
        xl = comDispatch('Excel.Application')
        xl.Workbooks.Open(patient_total_filename, ReadOnly=0)
        xl.Run("'" + macro_name + "'" + "!Patient_Total")
        xl.Workbooks(1).Close(SaveChanges=1)
        xl.Quit()
        del xl
        patient_total.configure(bg="green")



# main function
root = Tk()
w = 800 # width for the Tk root
h = 800 # height for the Tk root
# get screen width and height
ws = root.winfo_screenwidth() # width of the screen
hs = root.winfo_screenheight() # height of the screen
# calculate x and y coordinates for the Tk root window
x = (ws/1.1) - (w/1.1)
y = (hs/4) - (h/4)
# set the dimensions of the screen
# and where it is placed
root.geometry('%dx%d+%d+%d' % (w, h, x, y))


labelfont = ('arial', 15, 'bold')
up_mac = Button(height=12, text='Step 1: Upload Macro File Here!', font = labelfont, command=macro_function)
time_track = Button(height=6,
                    text='Step 1: Upload Time Tracking Results here. \n Step 2: Choose File Name and Location'
                             ' For Time Tracking Output Text File.', font = labelfont, command=time_tracking_function)
adenoma_path = Button(height=6,
                      text='Step 1: Upload Adenoma file from Path Lab here. \n Step 2: Choose File Name and Location'
                           ' For # of Adenomas Output Text File.', font=labelfont, command=adenoma_function)
patient_total = Button(height=10,
                      text='Step 1: Upload Total Patient File From Provation here. \n Step 2: Choose File Name and'
                           ' Location For # of Total Patients Output Text File.', font=labelfont, command=patient_total_function)



up_mac.pack(fill=BOTH)
time_track.pack(fill=BOTH)
adenoma_path.pack(fill=BOTH)
patient_total.pack(fill=BOTH)


up_mac.configure(bg = "yellow", fg = "black")
time_track.configure(bg = "red")
adenoma_path.configure(bg = "red")
patient_total.configure(bg = "red")


adenoma_path["state"] = DISABLED
time_track["state"] = DISABLED
patient_total["state"] = DISABLED


root.title("Ben Sklar's Adenoma Program!")
mainloop()