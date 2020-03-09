'''
Create date and time variable for batch script

Created by Shaun Thomas
Version 1.0

08/7/2017


1.  Check existence of all directories; 
    create directory if needed
        FTP - incoming and ack
        archive - processed, duplicate, nomatch, mailedfiles
        production - PRINT FILES
        log - log
    
2.  Open and read log. 
    Create list of filenames from the log (Check for Duplicates)

3.  Read each incoming file.
    Create lists for data records, header record and footer record
    Verify that the administrative information in the header and footer records match the data itself.
    (Check for Data Matching)

4.  Update log with "DUPLICATE" if file is duplicated in the log. 
    Move file to the "duplicate" folder.

5.  Update log with "NO MATCH" if the administrative information does not match the data. 
    Move file to the "nomatch" folder.

6.  Process the file for production if the admin info matches
    and the file is not in the log,
    
    a.  Copy file to the "processed" folder.
        Create folder for the current day in the "PRINT FILES"
        Move the file to the current day folder.

    b.  Create acknowledgement files for each file.
        Move them to the FTP "ack" folder.

    c.  Create the print files via PrintNet. 

    d.  Create and send email to production with list of files to process
    
    e.  Create and send email to VG and DOF with list of ack files available
    
  


'''

import os
import sys
import csv
import time
import datetime
import re
import shutil
import subprocess
import smtplib
from email.mime.text import MIMEText
import collections
from math import ceil
import win32com.client as win32


def main():
    
    # SET DIRECTORY VARIABLES 
    
    
    # Get current day and month to create current sub-folders
    currentyear = datetime.date.today().strftime('%Y')
    currentmonth = datetime.date.today().strftime('%Y%m')
    currentday = datetime.date.today().strftime('%Y%m%d')
    
    
    # Working mapped production drive
    def set_E_paths():
        """ Folder paths for E Drive """
        drive_path = r'E:\Share\NAS\unix\workflow\nycdof\daily\CMVT_NEW'
        printnet = "E:\GMC\PrintNet T Designer\PNetTC.exe"
        scripts_folder = os.path.join(drive_path, "prog", "scripts")
        workflow_path = os.path.join(drive_path, "prog", "PrintNet")
        ftp_ack_dir = r'F:\nycdofcn\BTSFLAT\ack'
        ftp_incoming_dir = r'F:\nycdofcn\BTSFLAT\incoming' 
        log_path = os.path.join(drive_path, "log")
        mail_test = False        
        return drive_path, printnet, scripts_folder, workflow_path, ftp_ack_dir, ftp_incoming_dir, log_path, mail_test
    
    def set_G_paths():
        """ Folder paths for G Drive """
        drive_path = r'X:\workflow\nycdof\daily\CMVT_NEW'
        printnet = "G:\PrintNet T Designer\PNetTC.exe"
        scripts_folder = os.path.join(drive_path, "prog", "scripts")
        workflow_path = os.path.join(drive_path, "prog", "PrintNet")
        ftp_ack_dir = r'F:\nycdofcn\BTSFLAT\ack'
        ftp_incoming_dir = r'F:\nycdofcn\BTSFLAT\incoming'
        mail_test = False       
        log_path = os.path.join(drive_path, "log")
        return drive_path, printnet, scripts_folder, workflow_path, ftp_ack_dir, ftp_incoming_dir, log_path, mail_test
    
    def set_test_paths(drive_paths):
        """ Folder paths for Testing. Take the drive_path, 
        printnet, scripts_folder, workflow_path from either 
        E of G drive """
        drive_path, printnet, scripts_folder, workflow_path = drive_paths[:4]
        test_drive = os.path.join(drive_path, "prog", "Testing")
        ftp_ack_dir = os.path.join(test_drive, "FTP", "ack")
        ftp_incoming_dir = os.path.join(test_drive, "FTP", "incoming")
        log_path = os.path.join(test_drive, "log")
        mail_test = True        
        return test_drive, printnet, scripts_folder, workflow_path, ftp_ack_dir, ftp_incoming_dir, log_path, mail_test
    
    
    ### Paths for PrintNet and Scripts 
    # drive_path, printnet, scripts_folder, workflow_path, ftp_ack_dir, ftp_incoming_dir, log_path, mail_test = set_E_paths()
    drive_path, printnet, scripts_folder, workflow_path, ftp_ack_dir, ftp_incoming_dir, log_path, mail_test = set_G_paths()
    # drive_path, printnet, scripts_folder, workflow_path, ftp_ack_dir, ftp_incoming_dir, log_path, mail_test = set_test_paths(set_E_paths())
    # drive_path, printnet, scripts_folder, workflow_path, ftp_ack_dir, ftp_incoming_dir, log_path, mail_test = set_test_paths(set_G_paths())
    
    
    email_path = r'X:\workflow\nycdof\daily\CMVT_NEW'

    # Archive directories
    archive_dir = os.path.join(drive_path, "archive")
    
    mailedfiles_dir = os.path.join(archive_dir, "mailedfiles")
    duplicate_dir = os.path.join(archive_dir, "duplicate")
    nomatch_dir = os.path.join(archive_dir, "nomatch")
    
    # Verify archive directories and create if needed
    for dir in [mailedfiles_dir, duplicate_dir, nomatch_dir]:
        if not os.path.exists(dir):
            os.makedirs(dir)

            
    # Initialize folders for current production data and print files
    print_dir = ""
    print_data_dir = ""
    print_prod_dir = ""
    email_print_dir = ""
    ack_dir = ""
    
    # Ensure incoming files directory exists
    if not os.path.exists(ftp_incoming_dir):
        print "Unable to find directory for incoming files."
        sys.exit()
     
    # Get list of files in incoming folder. Process if there are files
    incoming_data_tup_list = []
    for file in os.listdir(ftp_incoming_dir):
        ftp_incoming_file = os.path.join(ftp_incoming_dir, file)  # FTP site
        filenameMatch = re.match(r'^BTS_CMV_STK_', file)  # Naming convention for CMVT data    
        if os.path.isfile(ftp_incoming_file) and filenameMatch != None:
            incoming_data_tup_list.append((file, ftp_incoming_file))
    
    
    if len(incoming_data_tup_list) > 0: 
        
        # READ EACH FILE ON THE FTP SITE       
        dataFilesSpecsList = [] # List of data files and specifications for production email
        ackFilesList = []       # List of ack files for client email
        
        # Process each incoming file 
        for file, ftp_incoming_file in incoming_data_tup_list:
        
            # date_mod = os.path.getmtime(ftp_incoming_file)
            # dtup = datetime.datetime.fromtimestamp(date_mod).timetuple()
            # file_date_recvd = "{}{:0>2}{:0>2}".format(dtup.tm_year, dtup.tm_mon, dtup.tm_mday)
        
            
            file_day = currentday
            file_month = currentmonth
            file_year = currentyear
            
            # VERIFY AND/OR CREATE LOG FILE. GET LIST OF PREVIOUSLY PROCESSED FILES
            
            logManager = LogManager(log_path, file_year)
            logHeader = logManager.logHeader
            
            
            """ COMPARE EACH FILE TO THE LOG AND HEADER/FOOTER RECORDS OF DATA
                PROCESS ACCORDINGLY  """
            
            # Read data and get admin info and match status
            dataFileSpecs = DataRecordAndMailCounter(ftp_incoming_file, currentday)
            
            logUpdater = LogEntryCreater(dataFileSpecs, logHeader) 
            currentLogEntry = logUpdater.checkMatchStatus()

            # Check file status and process accordingly
            if file in logManager.getFilesInLog():
                currentLogEntry[logHeader.index("Status")] = "DUPLICATE"
                logManager.writeEntry(currentLogEntry)
                
                shutil.move(ftp_incoming_file, os.path.join(duplicate_dir, file))
                dataFilesSpecsList.append("{}:  DUPLICATE. NOT AVAILABLE FOR PROCESSING.".format(file))
                ackFilesList.append("{}:  DUPLICATE. NOT AVAILABLE FOR PROCESSING.".format(file))
                
            
            elif "N" in currentLogEntry[logHeader.index("First_Stamp_Match"):logHeader.index("Last_Plate_Match")+1]:
                currentLogEntry[logHeader.index("Status")] = "NO MATCH"
                logManager.writeEntry(currentLogEntry)
                shutil.move(ftp_incoming_file, os.path.join(nomatch_dir, file))
                
                unmatched_indices = [i for i in range(len(currentLogEntry)) if currentLogEntry[i] == "N"]
                unmatched_fields = [logHeader[n] for n in unmatched_indices]
                dataFilesSpecsList.append("{}:  NO MATCH - {}. NOT AVAILABLE FOR PROCESSING.".format(file, ", ".join(unmatched_fields)))
                ackFilesList.append("{}:  NO MATCH - {}. NOT AVAILABLE FOR PROCESSING.".format(file, ", ".join(unmatched_fields)))
                
            
            else:    

                file_date_recvd = dataFileSpecs.dateReceived
                file_date_processed = dataFileSpecs.dateProcessed
                file_month = file_date_processed[:6]
                file_year = file_date_processed[:4]
                
                
                # Folder for current production data and print files
                print_dir = os.path.join(drive_path, "PRINT_FILES", file_month, file_date_processed)
                print_data_dir = os.path.join(print_dir, "Data")
                print_file = os.path.join(print_data_dir, file)       # PRINT_FILES folder
                print_prod_dir = os.path.join(print_dir, "Print")
                email_print_dir = os.path.join(email_path, "PRINT_FILES", file_month, file_date_processed)
                ack_dir = os.path.join(archive_dir, "ackfiles", file_month, file_date_processed)
                
                
                # Verify and/or create daily archive and print folders  
                if not os.path.exists(print_dir):
                    os.makedirs(print_data_dir)
                    os.mkdir(print_prod_dir)
                
                if not os.path.exists(ack_dir):
                    os.makedirs(ack_dir)
                
                currentLogEntry[logHeader.index("Status")] = "PRINT"
                logManager.writeEntry(currentLogEntry)
                
                # Copy/Move original data to archive and print folders
                shutil.move(ftp_incoming_file, print_file)
                
                # Create ack file path and name
                filenamedate = file.split(".")[0].lstrip("BTS_CMV_STK_")
                ack_file = "CMV_ACK_{}.dat".format(filenamedate)
                ack_path = os.path.join(print_data_dir, ack_file)

                # Create acknowledgement files. Copy to FTP Ack folder and move to archive 
                create_ack_script = os.path.join(scripts_folder, "CreateAckFile.py")
                subprocess.call(["python", create_ack_script, print_file, ack_path])
                shutil.copy2(ack_path, os.path.join(ftp_ack_dir, ack_file))
                shutil.move(ack_path, os.path.join(ack_dir, ack_file))
                
                # Create print files
                workflow = os.path.join(workflow_path, "DOF_CMVT_Stickers.wfd")
                configfile = os.path.join(workflow_path, "CMVT.job")
                psfile = os.path.join(print_prod_dir, "{}.ps".format(file.split(".")[0]))
                seqfile = os.path.join(print_data_dir, "sequence_{}".format(file))
                logfile = os.path.join(print_data_dir, "gmc_{}.log".format(file_date_processed))
                
                out_module = dataFileSpecs.single_multi_counts["output"]
                
                subprocess.call([printnet, workflow,
                    "-difDataIn", print_file,
                    "-FileParams", file,
                    "-DateReceivedParams", file_date_recvd,
                    "-DataProcessedParams", file_date_processed,
                    "-o", out_module, 
                    "-f", psfile,
                    "-c", configfile,
                    "-pc", "OCE",
                    "-dc", "CMVT",
                    "-e", "AdobePostScript3",
                    "-la", logfile])
                    
                if dataFileSpecs.single_multi_counts["size"] == "SMALL":    
                    subprocess.call([printnet, workflow,
                        "-difDataIn", print_file,
                        "-FileParams", file,
                        "-DateReceivedParams", file_date_recvd,
                        "-DataProcessedParams", file_date_processed,
                        "-o", "SeqFile", 
                        "-f", seqfile,
                        "-datacodec", "UTF-8",
                        "-dataoutputtype", "CSV"]) 
                
                specsFomatted = DataSpecsFormatter(dataFileSpecs)
                specsFomatted = specsFomatted.format()
                print specsFomatted
                dataFilesSpecsList.append(specsFomatted)
                
                ackFilesList.append(ack_file)

                
        
        # Create and send email message with files and file specs
            
        prod_email = ProductionEmail(dataFilesSpecsList, email_print_dir, currentday, mail_test)            
        sendEmailMsgPY(
            prod_email.sender,
            prod_email.mailingList,
            prod_email.msgSubject(),
            prod_email.msgText())
        
        client_email = ClientEmail(ackFilesList, currentday, mail_test)
        sendEmailMsgOL(
            client_email.mailingList,
            client_email.msgSubject(),
            client_email.msgText())
    
    else:
        nofiles_email = NoFilesEmail(currentday)
        sendEmailMsgPY(
            nofiles_email.sender,
            nofiles_email.mailingList,
            nofiles_email.msgSubject(),
            nofiles_email.msgText())
    
    
class LogManager(object):
    """ Class to create, read and write to the log"""
    
    def __init__(self, log_path, year):
        self.log = os.path.join(log_path, "cmvt_{}.log".format(year))
        self.logHeader = ["Filename", "Sticker_Type", "Date_Received","Date_Processed",
            "Record_Count", "Stamp_Count", "Small_File_Count", "Large_File_Count",
            "Single_Sticker_Count", "Multiple_Sticker_Count", "Mail_Piece_Count",
            "Record_Count_Match", "Stamp_Count_Match", "First_Stamp_Match",
            "Last_Stamp_Match", "First_Plate_Match", "Last_Plate_Match",
            "Status", "Date Mailed", "Ack_Processed_By_DOF"]
        
        self.createLog()
    
    def createLog(self):
        """ Create new log for the year if it doesn't exist """
        if not os.path.exists(self.log):
            with open(self.log, 'wb') as w:
                logWriter = csv.writer(w, delimiter="|")
                logWriter.writerow(self.logHeader)

    def readEntries(self):
        """ Read log entries into csv file object """
        logEntries = []
        with open(self.log, 'rb') as r:
            logReader = csv.reader(r, delimiter="|")
            logReader.next()
            logEntries = [e for e in logReader]
        return logEntries

    def getFilesInLog(self):
        """ Create list of data files in the log """
        dataFilesInLog = []
        logEntries = self.readEntries()
        for entry in logEntries:
            dataFilename = entry[self.logHeader.index("Filename")]
            dataFilesInLog.append(dataFilename)
        return dataFilesInLog
        
    def writeEntry(self, entry):
        """ Append entry to the log """
        with open(self.log, 'ab') as a:
            logAppender = csv.writer(a, delimiter="|")
            logAppender.writerow(entry)
    
        
class DataRecordAndMailCounter(object):
    """ Class to parse each data file and get the Sticker 
    Color, Record counts, Mail Piece counts, Plate Numbers 
    and Stamp Numbers  """
    
    def __init__(self, file, currentday):                   
        
        # File metadata
        self.file = file
        self.filename = os.path.basename(file)
        self.file_type = self.filename.split('.')[0][-3:]  # SML or LRG
        self.dateReceived = self.getDateReceived()
        self.dateProcessed = currentday
        
        # Parse data
        self.data_dict = self.readDataToDict()
        self.recList = self.data_dict["1"]        
        self.hdrList = self.data_dict["HDR"][0]   
        self.ftrList = self.data_dict["FTR"][0]
        
        # Field names for each record type
        self.recfields = ["RecordType", "RowNumber", "PlateState", "PlateNo", "VehicleCode",
            "Weight", "PlateCharge", "PaymentDate", "StampNo", "BTSaccountNo", 
            "CustomerName", "Address1", "Address2", "City", "State", "Zip"]

        self.hdrfields = ["Header Designation", "File Name", "File Record Count",
            "File Creation Date", "Paper Color"]
        
        self.ftrfields = ["Footer Designation", "Total number of stamps in file",
            "First Stamp No", "First Plate", "Last Stamp No", "Last Plate"]
        
        # Info from records
        self.recordCount = len(self.recList)
        self.dataFirstStamp = self.recList[0][self.recfields.index("StampNo")]
        self.dataLastStamp = self.recList[-1][self.recfields.index("StampNo")]
        self.dataFirstPlate = self.recList[0][self.recfields.index("PlateNo")]
        self.dataLastPlate = self.recList[-1][self.recfields.index("PlateNo")]
        
        self.dataStampCount = self.getDataStampCount()
        
        self.AddressDict = self.createAddressDict()
        self.record_and_mail_counts_dict = self.create_Record_And_Mail_Counts_Dict()
        
        # Info from header / footer
        self.stampColor = self.hdrList[self.hdrfields.index("Paper Color")]
        self.hdrRecordCount = int(self.hdrList[self.hdrfields.index("File Record Count")])
        self.ftrStampCount = int(self.ftrList[self.ftrfields.index("Total number of stamps in file")])    
        self.ftrFirstStamp = self.ftrList[self.ftrfields.index("First Stamp No")]
        self.ftrLastStamp = self.ftrList[self.ftrfields.index("Last Stamp No")]   
        self.ftrFirstPlate = self.ftrList[self.ftrfields.index("First Plate")]
        self.ftrLastPlate = self.ftrList[self.ftrfields.index("Last Plate")]   
        
        self.single_multi_counts = self.getSingleAndMultiCounts()
        self.impression_count_dict = self.createPrintImpressionCountDict()
        
    def getDateReceived(self):
        """ Get the modified date of the file. 
            Same as the date received/sent to the FTP server. """
        
        date_mod = os.path.getmtime(self.file)
        dtup = datetime.datetime.fromtimestamp(date_mod).timetuple()
        date_mod_fmt = "{}{:0>2}{:0>2}".format(dtup.tm_year, dtup.tm_mon, dtup.tm_mday)
        return date_mod_fmt

    
    def readDataToDict(self):    
        """ Read each line in the data into a dictionary 
        with record type ("HDR", "FTR" or "1") as key. """
        
        data_dict = collections.defaultdict(list)
        with open(self.file, 'rb') as d:
            csvIn = csv.reader(d, delimiter='|')
            for line in csvIn:
                recordtype = line[0]
                data_dict[recordtype].append(line)
        return data_dict

        
    def getDataStampCount(self):
        """ Count number of unique stamps in file """
        
        stampList = []
        for row in self.recList:
            stamp = row[self.recfields.index("StampNo")] 
            stampList.append(stamp)
        return len(set(stampList))

        
    def createAddressDict(self):        
        """ Count the occurrences of an mailing address.
        Create dictionary with addresses and counts. """
        addressDict = collections.defaultdict(int)
        for row in self.recList:
            StreetAddress = row[self.recfields.index("Address1")].split()
            StreetAddressString = " ".join(StreetAddress) 
                
            CustomerName = row[self.recfields.index("CustomerName")].split()
            CustomerNameString = " ".join(CustomerName) 
            NameFirstChars = CustomerNameString[:12] if len(CustomerNameString) >= 12 else CustomerNameString
                
            Zip5 = row[self.recfields.index("Zip")]
            Zip5 = Zip5[:5] if len(Zip5) >= 5 else Zip5 
            
            address_string = StreetAddressString + NameFirstChars + Zip5
            addressDict[address_string] += 1
        return addressDict
        
    
    def create_Record_And_Mail_Counts_Dict(self):        
        """ Take dict with Address counts and count the 
        occurrences of the address counts.
        
        Create a new dict with the number of Stickers per Record 
        as the key and a list with the number of records 
        and number of mail pieces as the values. 
        
        Envelopes fit up to 20 stickers. For sticker counts above t
        """
         
        addressCounts = self.AddressDict.values()
        stkrsPerRecCounter = collections.Counter(addressCounts)
        stkrsPerRecDict = dict(stkrsPerRecCounter)
        
        for stk_per_rec in stkrsPerRecDict.keys():
                
            rec_count = stkrsPerRecDict.get(stk_per_rec)
            mail_count = stkrsPerRecDict.get(stk_per_rec)
            stkrsPerRecDict[stk_per_rec] = [rec_count, mail_count]
            
            if stk_per_rec > 20:
                envs_per_rec = int(ceil(stk_per_rec/20.0))
                new_mail_count = envs_per_rec * rec_count
                stkrsPerRecDict[stk_per_rec][1] = new_mail_count                    
        
        return stkrsPerRecDict
        
        
    def createPrintImpressionCountDict(self):
        """ Count the print impressions for the file. 
        Calculate the number of cover sheets, 
        separator sheets and sticker sheets used. 
        Round up the sticker sheet counts to get the last sheet. """
        
        # SML and LRG files have a cover sheet
        cover_sheet_count = 1
        
        separator_sheets_count = 0
        stkr_imp_count = 0
        
        if self.file_type == "SML":
            # Each count group has a separator sheet
            separator_sheets_count = len(self.record_and_mail_counts_dict.keys())
            
            # Calculate the 3UP sticker sheets for each count group. 
            for k in self.record_and_mail_counts_dict.keys():
                stk_per_rec = k
                rec_count = self.record_and_mail_counts_dict[stk_per_rec][0]
                group_stk_count = stk_per_rec * rec_count
                group_imp_count = int(ceil(group_stk_count/3.0))
                stkr_imp_count += group_imp_count
        else:
            separator_sheets_count = 1
            stkr_imp_count = int(ceil(self.recordCount/3.0))

        return {"color": cover_sheet_count + separator_sheets_count,
                "sticker": stkr_imp_count,
                "total": cover_sheet_count + separator_sheets_count + stkr_imp_count}
        
        
    def getSingleAndMultiCounts(self):
        """ Full name of sticker file type. To be recorded in the file log. """
        
        single_count = ""
        if self.record_and_mail_counts_dict.get(1) == None:
            single_count = 0     
        else:
            single_count = self.record_and_mail_counts_dict[1][0]
        
        multi_count = ""
        if self.record_and_mail_counts_dict.get(1) == None:
            multi_count = self.recordCount 
        else:
            multi_count = self.recordCount - self.record_and_mail_counts_dict[1][0]
        
        mail_count = sum(i[1] for i in self.record_and_mail_counts_dict.values())
        
        type_dict = {
        "SML" : {
                "size" : "SMALL",
                "mail_msg" : "SMALL FILE - INSERT AND MAIL",
                "small_count" : self.recordCount,
                "large_count" : 0,
                "singles_count" : single_count,
                "multi_count" : multi_count, 
                "mail_count" : mail_count,
                "output" : "Sticker_SML_3UP"
                },
        
        "LRG" : {
                "size" : "LARGE",
                "mail_msg" : "LARGE FILE - RETURN TO DOF",
                "small_count" : 0,
                "large_count" : self.recordCount,
                "singles_count" : 0,
                "multi_count" : 0,
                "mail_count" : 0,
                "output" : "Sticker_LRG_3UP"
                }
        }
        return type_dict[self.file_type]


    def RecordAndMailCountsList(self):
        """ Create list of records and mail piece counts.
            Counts are displayed in email to production. """
        
        stickerCountsList = []
        for k in sorted(self.record_and_mail_counts_dict.keys()): 
            record_count = self.record_and_mail_counts_dict[k][0]
            mail_piece_count = self.record_and_mail_counts_dict[k][1]
            stickerCountsList.append((k, record_count, mail_piece_count))
        return stickerCountsList

    
    
class LogEntryCreater(object):
    def __init__(self, dataSpecs, logHeader):
        self.dataSpecs = dataSpecs
        self.logHeader = logHeader
        
        self.log_entry = [self.dataSpecs.filename,
            self.dataSpecs.single_multi_counts["size"], 
            self.dataSpecs.dateReceived,
            self.dataSpecs.dateProcessed,
            self.dataSpecs.recordCount, 
            self.dataSpecs.dataStampCount, 
            self.dataSpecs.single_multi_counts["small_count"], 
            self.dataSpecs.single_multi_counts["large_count"], 
            self.dataSpecs.single_multi_counts["singles_count"], 
            self.dataSpecs.single_multi_counts["multi_count"], 
            self.dataSpecs.single_multi_counts["mail_count"], 
            "N", "N", "N", "N", "N", "N", "","", "N"]
    
    def checkMatchStatus(self):
        """ Verify that the values from the data match the 
        values in the header and footer records. Returns log 
        entry as a list. """
        
        if self.dataSpecs.recordCount == self.dataSpecs.hdrRecordCount:
            self.log_entry[self.logHeader.index("Record_Count_Match")] = "Y"
        
        if self.dataSpecs.getDataStampCount() == self.dataSpecs.ftrStampCount:    
            self.log_entry[self.logHeader.index("Stamp_Count_Match")] = "Y"
        
        if self.dataSpecs.dataFirstStamp == self.dataSpecs.ftrFirstStamp:
            self.log_entry[self.logHeader.index("First_Stamp_Match")] = "Y"
        
        if self.dataSpecs.dataLastStamp == self.dataSpecs.ftrLastStamp:    
            self.log_entry[self.logHeader.index("Last_Stamp_Match")] = "Y"

        if self.dataSpecs.dataFirstPlate == self.dataSpecs.ftrFirstPlate:
            self.log_entry[self.logHeader.index("First_Plate_Match")] = "Y"
        
        if self.dataSpecs.dataLastPlate == self.dataSpecs.ftrLastPlate:    
            self.log_entry[self.logHeader.index("Last_Plate_Match")] = "Y"
        
        return self.log_entry
  
  
class DataSpecsFormatter(object):
    def __init__(self, dataSpecs):
        self.filename = dataSpecs.filename
        self.mail_type = dataSpecs.single_multi_counts["size"]
        self.mail_msg = dataSpecs.single_multi_counts["mail_msg"]
        
        self.recordCount = dataSpecs.recordCount
        self.mail_count = dataSpecs.single_multi_counts["mail_count"]
        self.RecordAndMailCountsList = dataSpecs.RecordAndMailCountsList()
        
        self.stampColor = dataSpecs.stampColor
        self.dataFirstStamp = dataSpecs.dataFirstStamp
        self.dataFirstPlate = dataSpecs.dataFirstPlate
        self.dataLastStamp = dataSpecs.dataLastStamp
        self.dataLastPlate = dataSpecs.dataLastPlate
        self.impression_count_dict = dataSpecs.impression_count_dict

    
    def format(self):
        format_type = {
                       "SMALL": self.createSmallFormat(),
                       "LARGE": self.createLargeFormat()
                       }
        return format_type[self.mail_type]
        
 
    def createSmallFormat(self):    
        """ Return a string with record counts, single and multi 
        sticker counts  and first and last stamp and plate numbers 
        to include in Production email. """
        
        countList = []
        for group, rec_cnt, mail_cnt in self.RecordAndMailCountsList:
            if group > 20: 
                countList.append("{}      :      {}  ({} mail pieces)".format(group, rec_cnt, mail_cnt))
            else:
                countList.append("{}      :      {}".format(group, rec_cnt))

        totalSheetCount   = self.impression_count_dict["total"]
        colorSheetCount   = self.impression_count_dict["color"]
        stickerSheetCount = self.impression_count_dict["sticker"]
        
        return "\r\n".join([
            "{:<50}{}".format("FILE: " + self.filename, self.mail_msg),
            "",
            "Sticker Color: {}".format(self.stampColor),
            "",
            "Stickers: {}             Mail Pieces: {}".format(self.recordCount, self.mail_count),
            "",
            "Stickers          :    Record/Mail",
            "Per Recipient          Counts",
            "*******************************************",
            "\r\n".join(countList),
            "",
            "Print Impressions: {}".format(totalSheetCount),
            "Color sheets: {}      Sticker sheets: {}".format(colorSheetCount, stickerSheetCount),
            "", 
            "First Stamp No: {}           First Plate No: {}".format(self.dataFirstStamp, self.dataFirstPlate),
            "Last Stamp No: {}            Last Plate No: {}".format(self.dataLastStamp, self.dataLastPlate),
            ""
            ])
        
        
    def createLargeFormat(self):
        """ Return a string with record counts, single and multi 
        sticker counts  and first and last stamp and plate numbers 
        to include in Production email. """
        
        totalSheetCount   = self.impression_count_dict["total"]
        colorSheetCount   = self.impression_count_dict["color"]
        stickerSheetCount = self.impression_count_dict["sticker"]
        
        return "\r\n".join([
            "{:<50}{}".format("FILE: " + self.filename, self.mail_msg),
            "",
            "Sticker Color: {}".format(self.stampColor),
            "",
            "Records: {}".format(self.recordCount),
            "",
            "Print Impressions: {}".format(totalSheetCount),
            "Color sheets: {}      Sticker sheets: {}".format(colorSheetCount,stickerSheetCount),
            "", 
            "First Stamp No: {}           First Plate No: {}".format(self.dataFirstStamp, self.dataFirstPlate),
            "Last Stamp No: {}            Last Plate No: {}".format(self.dataLastStamp, self.dataLastPlate),
            "" 
            ])
        
        
def sendEmailMsgPY(sender, mailingList, subject, msgText):
    """ Function to send Email using smtp module. 
    Use to send emails to addresses within AST network.  
    Uses dsproduction@amstock.com as sender"""
    
    ''' NOTE: Mail3.amstock.com decommissioned Jan 2020 
        Switched function with 'sendEmailMsgOL' 
        
        NOTE: AST uses Mail.amstock.com for to 
        send emails for its automated process. FEB 2020'''
    
    msg = MIMEText(msgText)
    msg['From'] = sender
    msg['To'] = ",".join(mailingList)
    msg['Subject'] = subject

    try:
        # Send the email via AST SMTP server.
        server = smtplib.SMTP("Mail.amstock.com")
        recipients_not_mailed = server.sendmail(sender, mailingList, msg.as_string())
        server.quit()
    
        if len(recipients_not_mailed) == 0:
            print "All recipients were successfully contacted"
        else:
            print "Following recipients were rejected: {}".format(", ".join(recipients_mailed.keys()))
    
    except SMTPSenderRefused as sr:
        print sr
    except SMTPRecipientsRefused as rr:
        print rr 
    except SMTPDataError as de:
        print de
    except SMTPException as se:
        print se
    except:
        print "Unknown error. Unable to send email"

        
def sendEmailMsgOL(mailingList, msgSubject, msgText):
    """ Function to send Email using smtp module. 
    Use to send emails to clients outside of AST network. 
    Uses personal AST email as sender"""
    
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = ";".join(mailingList)
    mail.Subject = msgSubject
    mail.Body = msgText
    mail.Send()


    
class ProductionEmail(object):
    """ Email send to DS production team """
    
    def __init__(self, DataFilesSpecsList, print_prod_dir, currentday, mail_test):
        self.DataFilesSpecsList = DataFilesSpecsList
        self.print_prod_dir = print_prod_dir
        self.currentday = currentday
        self.sender = "dsproduction@astfinancial.com"        
        self.mailingList = self.createMailingList(mail_test)       
        
    def createMailingList(self, mail_test):
        if mail_test == True:
            return ["sthomas@astfinancial.com"]
        else:
            return ["kmcneil@astfinancial.com","dschwarz@astfinancial.com",
            "sthomas@astfinancial.com",
            "svogt@astfinancial.com","jherrera@astfinancial.com",
            "norellana@astfinancial.com","lprice@astfinancial.com",
            "rchetram@astfinancial.com","yjainarain@astfinancial.com"]
            
            
    def msgSubject(self):
        """ Email subject line """
        
        return "DOF CMVT Stickers - Print Files - {}".format(self.currentday)
        
       
    def msgText(self):   
        """Combine Mailing date and list of files with counts 
        into single message. Convert message to string."""
        
        divider = "\r\n\r\n{0}\r\n{0}\r\n\r\n".format("-" * 70)
        
        return "\r\n".join([
            "Files below are available at :",
            "",
            self.print_prod_dir,
            "",
            "",
            divider.join(self.DataFilesSpecsList)
            ])

        
class ClientEmail(object):
    def __init__(self, ackFilesList, currentday, mail_test):
        self.ackFilesList = ackFilesList
        self.currentday = currentday
        # self.sender = "dsproduction@astfinancial.com"        
        self.mailingList = self.createMailingList(mail_test)      
    
    def createMailingList(self, mail_test):
        if mail_test == True:
            return ["sthomas@astfinancial.com"]
        else:
            return ["sthomas@astfinancial.com", "WashingtonS@finance.nyc.gov",
            "MalatestaA@finance.nyc.gov", "HeywardR@finance.nyc.gov",
            "ZimmermanL@finance.nyc.gov","abhagwandin@vanguarddirect.com",
            "mmuniz@vanguarddirect.com","kswan@hellovanguard.com",
            "Team3@hellovanguard.com"]
    
    def msgSubject(self):
        return "DOF CMVT Stickers - ACK Files - {}".format(self.currentday)
        
       
    def msgText(self):   
        """Combine Mailing date and list of files with counts 
           into single message. Convert message to string."""
        
        return "\r\n".join([
            "The 'ack' files below are available on the AST DS FTP Site:",
            "",
            "\r\n\r\n".join(self.ackFilesList)
            ])      


class NoFilesEmail(object):
    def __init__(self, currentday):
        self.currentday = currentday
        self.sender = "dsproduction@astfinancial.com"        
        self.mailingList = ["sthomas@astfinancial.com"]       
    
    def msgSubject(self):
        return "DOF CMVT Stickers - NO PRINT FILES - {}".format(self.currentday)
       
    def msgText(self):   
        return "No files available for processing for {}".format(self.currentday)

          
    
if __name__=="__main__":
    main()
