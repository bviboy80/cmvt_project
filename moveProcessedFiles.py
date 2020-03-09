import sys
import os
import csv
import datetime
import re
import collections


currentday = datetime.date.today().strftime('%Y%m%d')
currentmonth = datetime.date.today().strftime('%Y%m%d')
currentyear = datetime.date.today().strftime('%Y%m%d')


# FTP locations
ack_FTP_folder = r'F:\nycdofcn\BTSFLAT\ack'
ack_FTP_filelist = os.listdir(ack_FTP_folder)

mailed_FTP_folder = r'F:\nycdofcn\BTSFLAT\mailed'
mailed_FTP_filelist = os.listdir(mailed_FTP)


# Archive locations
# working_dir = r'E:\Share\NAS\unix\workflow\nycdof\daily\CMVT_NEW'
working_dir = r'X:\workflow\nycdof\daily\CMVT_NEW'
ack_archive = os.path.join(working_dir, 'archive', 'ackfiles')
mailed_archive = os.path.join(working_dir, 'archive', 'mailedfiles')


# Log Location
log_directory = os.path.join(working_dir, 'log')
log_directory_list = os.listdir(log_directory)

for log in log_directory_list:
    
    logFile = os.path.join(log_directory, log)
    
    stickerLog = StickerDataLog(logFile)              
                
    for filename in ack_FTP_filelist:
        ack_file = os.path.join(ack_FTP, filename)
            if os.path.isfile(ack_file) and file[-9:].upper() == "PROCESSED":
                
                archived_ack_file = ftp_ackfile.rstrip('.PROCESSED')
                archiveFolder = os.path.join(ack_archive, recvd_month, recvd_date) 
                try:
                    stickerLog.updateEntry(file)
                    recvd_date = stickerLog.getdateFileReceived(file)
                    recvd_month = received_date[:6]
                
                                      
                    shutil.move(ackfile, os.path.join(archiveFolder, file))

    stickerLog.writeLog()







class StickerDataLog(object):
    
    def __init__(self, logfile):
        self.logfile = logfile
        self.logHeader = ["Filename", "Date_Received", "Record_Count", "Stamp_Count",
            "Single_Count", "Multi_Count", "Record_Count_Match", "Stamp_Count_Match",
            "First_Stamp_Match","Last_Stamp_Match", "First_Plate_Match",
            "Last_Plate_Match", "Status", "Date Mailed", "Ack_Processed_By_DOF"]
    
        self.LogEntries = self.getLogEntries()
    
    def getLogEntries(self):
        logEntries = []
        if os.path.isfile(self.logfile):
            with open(self.logfile, 'rb') as o:
                logReader = csv.reader(o, delimiter="|")
                logReader.next()
                logEntries = [entry for entry in logReader]
        return logEntries
   
    def updateEntry(self, processed_ackfile):
        ackfile_stem = processed_ackfile[8:30]
        stickerfile = "BTS_CMV_STK_" + ackfile_stem
        for row in self.LogEntries:
            row[self.logHeader.index("Filename")] = logged_filename
            if stickerfile == logged_filename:
                row[self.logHeader.index("Ack_Processed_By_DOF")] = "Y"
                break
            else:
                continue
        
    def getdateFileReceived(self, processed_ackfile)
        ackfile_stem = processed_ackfile[8:30]
        stickerfile = "BTS_CMV_STK_" + ackfile_stem
        for row in self.LogEntries:
            row[self.logHeader.index("Filename")] = logged_filename
            if stickerfile == logged_filename:
                date_received = row[self.logHeader.index("Date_Received")]
                return date_received
                break
            else:
                continue
                
    def writeLog(self):
        with open(self.logfile, 'wb') as w:
            logWriter = csv.writer(w, delimiter="|")
            logWriter.writerow(self.logHeader)
            for entry in self.LogEntries:
                logWriter.writerow(entry)

            
            

        
class MailedDataLog(object):
    def __init__(self, logfile):
        self.logfile = logfile
        self.logHeader = ["Filename", "Date_Received", "Record_Count", "Stamp_Count",
            "Single_Count", "Multi_Count", "Record_Count_Match", "Stamp_Count_Match",
            "First_Stamp_Match","Last_Stamp_Match", "First_Plate_Match",
            "Last_Plate_Match", "Status", "Date Mailed", "Ack_Processed_By_DOF"]
    
        self.LogEntries = self.getLogEntries()
    
    def getLogEntries(self):
        logEntries = []
        if os.path.isfile(self.logfile):
            with open(self.logfile, 'rb') as o:
                logReader = csv.reader(o, delimiter="|")
                logReader.next()
                logEntries = [entry for entry in logReader]
        return logEntries
   
    def updateEntry(self, processed_ackfile):
        ackfile_stem = processed_ackfile[8:30]
        stickerfile = "BTS_CMV_STK_" + ackfile_stem
        for row in self.LogEntries:
            row[self.logHeader.index("Filename")] = logged_filename
            if stickerfile == logged_filename:
                row[self.logHeader.index("Ack_Processed_By_DOF")] = "Y"
                break
            else:
                continue
        
    def getdateFileReceived(self, processed_ackfile)
        ackfile_stem = processed_ackfile[8:30]
        stickerfile = "BTS_CMV_STK_" + ackfile_stem
        for row in self.LogEntries:
            row[self.logHeader.index("Filename")] = logged_filename
            if stickerfile == logged_filename:
                date_received = row[self.logHeader.index("Date_Received")]
                return date_received
                break
            else:
                continue
                
    def writeLog(self):
         with open(self.logfile, 'wb') as w:
            logWriter = csv.writer(w, delimiter="|")
            logWriter.writerow(self.logHeader)
            for entry in self.LogEntries:
                logWriter.writerow(entry)
        self.mailedfile
        
        
        
        

def moveFiles(ftp_folder, archive_folder):
    
    fileList = os.listdir(ftp_folder)
    for file in fileList:
        file_ext = file.split('.')[-1]
        if file_ext.upper() == "PROCESSED"
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
