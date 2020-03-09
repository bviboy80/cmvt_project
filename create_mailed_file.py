"""

Interface for selecting files that have been mailed daily
for CMVT stickers.

The incoming files log will be compared to the mailed files log, 
creating a list of files remaining to be mailed.

Resulting list will be displayed in the interface.

Files mailed for the day will be chosen from this list/or scanned in

compare the value to both list
if value in incoming but not in mailed - shift to mailed side (with option to shift back)

if not in either list, print in dialog box, "Please select an appropriate value"

"""

from Tkinter import *
import ttk
import os
import sys
import csv
import re
import collections
import shutil
import datetime
import smtplib
from email.mime.text import MIMEText
import win32com.client as win32

currentday = datetime.date.today().strftime('%Y%m%d')

work_dir= r'X:\workflow\nycdof\daily\CMVT_NEW'
# work_dir= r'E:\Share\NAS\unix\workflow\nycdof\daily\CMVT_NEW'
# work_dir= r'X:\workflow\nycdof\daily\CMVT_NEW\prog\Testing'


# Path with log files
log_directory = os.path.join(work_dir, 'log')
log_list = os.listdir(log_directory)


printfile_dir = os.path.join(work_dir, 'PRINT_FILES')
mailedfile_dir = os.path.join(work_dir, 'archive', 'mailedfiles')

mailedFTP = r'F:\nycdofcn\BTSFLAT\mailed'
# mailedFTP = r'X:\workflow\nycdof\daily\CMVT_NEW\prog\Testing\FTP\mailed'


# mailingList = ["sthomas@astfinancial.com"] 
       
mailingList = ["sthomas@astfinancial.com", "WashingtonS@finance.nyc.gov",
               "MalatestaA@finance.nyc.gov", "HeywardR@finance.nyc.gov",
               "ZimmermanL@finance.nyc.gov", "team3@hellovanguard.com",
               "abhagwandin@vanguarddirect.com", "mmuniz@vanguarddirect.com",
               "kswan@hellovanguard.com"] 



mailedList = []
notMailedList = []
ToBeMailedList = []
fileDict = collections.defaultdict(list)


# Read each log in the log folder
for file in log_list:
    logfile = (os.path.join(log_directory, file))
    if os.path.isfile(logfile) and logfile[-3:] == "log":
        with open(logfile, 'rb') as incoming:
            logReader = csv.reader(incoming, delimiter="|")
            logReader.next()
            
            # In each log, check for files that have not been mailed 
            for line in logReader:
                file_name = line[0]
                date_processed = line[3]
                month_processed = line[3][:6]
                mailedStatus = line[-2]
                
                if mailedStatus != "":
                    mailedList.append(line[0])
                else:
                    notMailedList.append(line[0])
                    fileDict[file_name] = [month_processed, date_processed]
                    
                    
# Set up base
root = Tk() 
root.title("CMVT Stickers - Mailed Files")

#Set Styles
s = ttk.Style()
s.theme_use('clam')
s.configure('Name.TLabel', foreground='white', background='blue', font="none 15 bold")
s.configure('Submit.TButton', foreground= 'black', background='green', font="none 8 bold")
s.configure('Reset.TButton', foreground='red', font="none 8 bold")

window = ttk.Frame(root, padding="10 10 12 12")
window.grid(column=0, row=0, sticky=(N, W, E, S))
window.columnconfigure(0, weight=1)
window.rowconfigure(0, weight=1)



# functions

# START CREATE MAILED FILE
    
def create_mailed_file():
    mailed_date = maileddate.get()
    mailedMonth = mailed_date[:6]
    mailedDir = os.path.join(mailedfile_dir, mailedMonth)
    if not os.path.exists(mailedDir):
        os.makedirs(mailedDir)
    
    outfile = "CMV_MAL_{}.txt".format(mailed_date)
    mailedFile = os.path.join(mailedDir, outfile)
    mailedFTPfile = os.path.join(mailedFTP, outfile)
    
    combinedData = combinedMailData()     # Create mailed file data
    if len(combinedData) > 0:
        with open(mailedFile, 'wb') as mf:
            recordCount = len(combinedData)
            
            csvOut = csv.writer(mf, delimiter='|')
            csvOut.writerow(["HDR", mailed_date, recordCount])
            for line in combinedData:
                csvOut.writerow(line)
            csvOut.writerow(["FTR", mailed_date, recordCount])
        
        shutil.copy(mailedFile, mailedFTPfile)
        
        updatelogs(mailed_date)
        
        mailedfile_email = MailedFilesEmail(outfile, currentday, mailingList)
        
        sendEmailOL(mailedfile_email.mailingList,
                    mailedfile_email.msgSubject(),
                    mailedfile_email.msgText())
                                        
    else:
        message.set('*** NO RECORDS. NO MAIL FILE CREATED. ***')
        
    
def combinedMailData():
    combinedDataList = []
    
    datafields = ["RecordType","RowNumber","PlateState","PlateNo","VehicleCode",
                  "Weight","PlateCharge","PaymentDate","StampNo","BTSaccountNo",
                  "CustomerName","Address1","Address2","City","State","Zip"]
    
              
    for filename in ToBeMailedList:
        processMonth = fileDict[filename][0]
        processDate = fileDict[filename][1]
        
        file = os.path.join(printfile_dir, processMonth, processDate, 'Data', filename)
        
        print file
        with open(file, "rb") as f: 
            csv_readfile = csv.reader(f, delimiter='|')
            
            for line in csv_readfile:
                if line[0]=="1":
                    combinedDataList.append(["1",
                    line[datafields.index("RowNumber")],
                    line[datafields.index("StampNo")],
                    line[datafields.index("PlateState")],
                    line[datafields.index("PlateNo")],
                    line[datafields.index("CustomerName")],
                    line[datafields.index("Address1")],
                    line[datafields.index("Address2")],
                    line[datafields.index("City")],
                    line[datafields.index("State")],
                    line[datafields.index("Zip")]])
    return combinedDataList
    
    
### UPDATE THE LOG WITH MAILED DATE
def updatelogs(mailed_date):
    for file in log_list:
        logData = []
        logfile = (os.path.join(log_directory, file))
        
        if os.path.isfile(logfile):
            with open(logfile, 'rb') as r:
                logReader = csv.reader(r, delimiter="|")
                logData = [line for line in logReader]
                
                for row in logData:
                    for filename in ToBeMailedList:
                        if filename == row[0]:
                            row[-2] = mailed_date
            with open(logfile, 'wb') as w:
                csvWriteLog = csv.writer(w, delimiter='|')
                for lines in logData:
                    csvWriteLog.writerow(lines)
    
    message.set('Mail file created for {}'.format(mailed_date))

    
def sendEmailOL(mailingList, msgSubject, msgText):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = ";".join(mailingList)
    mail.Subject = msgSubject
    mail.Body = msgText
    mail.Send()                            
                            
                            
class MailedFilesEmail(object):
    
    def __init__(self, outfile, currentday, mailingList):
        self.outfile = outfile
        self.currentday = currentday
        self.sender = "dsproduction@astfinancial.com"        
        self.mailingList = mailingList      
    
    def msgSubject(self):
        return "DOF CMVT Stickers - MAILED FILE - {}".format(self.currentday)
       
    def msgText(self):   
        """Combine Mailing date and list of files with counts 
           into single message. Convert message to string."""
        messageText = "The 'mailed' file below is available on the AST DS FTP Site:\r\n" + \
                      "\r\n\r\n" + self.outfile
        return messageText    
    
# END CREATE MAILED FILE
    
def confirm():
    pass

def confirm_yes():
    pass
    
def confirm_no():
    pass
    
def exit_window():
    pass

def addToTextEntry(event):
    w = event.widget
    sel = w.get(w.curselection())
    print sel
    
    
    
### UPDATE LISTS    
def add_mailed_file():
    switch_list_member(notMailedList, not_mailed_listbox,
                       ToBeMailedList, mailed_file_listbox)
 
 
def remove_mailed_file():
    switch_list_member(ToBeMailedList, mailed_file_listbox,
                       notMailedList, not_mailed_listbox)
 
 
def switch_list_member(list1, listbox1, list2, listbox2):
    """Get sticker name from Text Entry and 
    switch name from one list to the other."""
    
    sticker = sticker_name.get()
    
    if sticker in list1 and sticker not in list2:    
        list1.remove(sticker)
        update_list(list1, listbox1)
        list2.append(sticker)
        update_list(list2, listbox2)        
        sticker_input.delete(first=0, last=len(sticker))
        
        if list1 is notMailedList:
            message.set('ADDED:  {}'.format(sticker))
        elif list1 is ToBeMailedList: 
            message.set('REMOVED:  {}'.format(sticker))
    elif sticker in list2:
        message.set('{} ALREADY MOVED'.format(sticker))
    else:
        message.set('{} NOT FOUND!'.format(sticker))
    # print_lists()

    
def update_list(list, listbox):
    """Clear listbox content with insert new sorted list."""
    
    listbox.delete(0, "end")
    list.sort()
    for item in list:
        listbox.insert("end", item)
    
    
def reset():          
    sticker = sticker_name.get()
    sticker_input.delete(first=0, last=len(sticker))
    
    for item in ToBeMailedList:
        notMailedList.append(item)
    notMailedList.sort()
    del ToBeMailedList[:]
    
    update_list(not_mailed_listbox, not_mailed_listbox)
    update_list(mailed_file_listbox, mailed_file_listbox)
    
    message.set('RESET')

    
### END UPDATE LISTS         



            
                    
### VARIABLE TEXT VALUES
message = StringVar()
sticker_name = StringVar()
maileddate = StringVar()



### BUTTONS AND LABELS
   
# Title Label
name_label = ttk.Label(window, text="CMVT Sticker\n Mailed Files", style='Name.TLabel')
name_label.grid(row=1, column=2, sticky=N+W, columnspan=3)

# Reset
reset_button = ttk.Button(window, text="RESET", width=8, style='Reset.TButton', command=reset)
reset_button.grid(row=1, column=0, sticky=W, pady=10)

# Sticker Entry
sticker_input = ttk.Entry(window, width=40, textvariable=sticker_name)
sticker_input.grid(row=4, column=0, sticky=W, pady=20)

# Mailed Date Entry
date_input = ttk.Entry(window, width=20, textvariable=maileddate)
date_input.grid(row=4, column=3, sticky=W, padx=20)


# Files not yet Mailed
y_scrollLeft = ttk.Scrollbar(window, orient=VERTICAL)
y_scrollLeft.grid(row=5, column=1, sticky=N+S+W, rowspan=5)

not_mailed_listbox = Listbox(window, width=40, selectmode=EXTENDED, yscrollcommand=y_scrollLeft.set)
not_mailed_listbox.grid(row=5, column=0, sticky=E, rowspan=5)
y_scrollLeft['command'] = not_mailed_listbox.yview
not_mailed_listbox.bind("<Double-Button-1>", addToTextEntry)


# "Mailed Files" list
y_scrollRight = ttk.Scrollbar(window, orient=VERTICAL)
y_scrollRight.grid(row=5, column=4, sticky=N+S+W, rowspan=5)

mailed_file_listbox = Listbox(window, width=40, selectmode=EXTENDED, yscrollcommand=y_scrollRight.set)
mailed_file_listbox.grid(row=5, column=3, sticky=W, rowspan=5)
y_scrollRight['command'] = mailed_file_listbox.yview
mailed_file_listbox.bind("<Double-Button-1>", addToTextEntry)


# Add file to "Mailed Files" list
add_button = ttk.Button(window, text="Add >>>", width=12, command=add_mailed_file)
add_button.grid(row=5, column=2, sticky=NW, padx=20, pady=10)

# Remove file from "Mailed Files" list
remove_button = ttk.Button(window, text="<<< Remove", width=12, command=remove_mailed_file)
remove_button.grid(row=6, column=2, sticky=NW, padx=20)

# Create "Mailed Files" file
submit_button = ttk.Button(window, text="SUBMIT", width=12, style='Submit.TButton', command=create_mailed_file)
submit_button.grid(row=7, column=2, sticky=NW, padx=20)

# Message Label
message_label = ttk.Label(window, textvariable=message)
message_label.grid(row=1, column=3, columnspan=2, sticky=N, padx=20)

# Exit program
exit_button = ttk.Button(window, text="EXIT", width=12, command=exit_window)
exit_button.grid(row=8, column=2,sticky=NW, padx=20)

### END BUTTONS AND LABELS


# for child in window.winfo_children():
    # print child
    # print type(child)
    
    # child.grid_configure(padx=5, pady=5)

### INITIALIZE PENDING LISTBOX
notMailedList.sort()
for line in notMailedList:
    not_mailed_listbox.insert("end", line)
   
    
    
root.mainloop()
