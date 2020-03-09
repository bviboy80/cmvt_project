import os
import sys
import csv
import re

def main(data, outputfile):
    inFile = os.path.abspath(data)
    basefile = os.path.basename(inFile)
    
    headerRecord = ["RecordType","FileName","StampCount","FileDate","PantoneColor"]
    
    footerRecord = ["RecordType","StampsInFile","FirstStampNo",
                    "FirstPlateNo","LastStampNo","LastPlateNo"]
    
    dataRecord = ["RecordType","RowNumber","PlateState","PlateNo","VehicleCode",
                  "Weight","PlateCharge","PaymentDate","StampNo","BTSaccountNo",
                  "CustomerName","Address1","Address2","City","State","Zip"]

                  
    with open(inFile, 'rb') as In:
        with open(os.path.abspath(outputfile), 'wb') as ackOut:
            csvIn = csv.reader(In, delimiter='|')
            csvOut = csv.writer(ackOut, delimiter='|')
            
            headList, footList, dataList = [], [], []
            recordCount = 0
            
            for line in csvIn:
                dataLine = []
                if line[0] == "1":
                    
                    # Create list of data records and count records
                    recordCount+=1
                    dataLine.extend(["1",
                                     line[dataRecord.index("RowNumber")],
                                     line[dataRecord.index("StampNo")],
                                     line[dataRecord.index("PlateState")],
                                     line[dataRecord.index("PlateNo")]])
                    dataList.append(dataLine)
                
                # Create header and footer. Make place holder for record counts.
                elif line[0] == "HDR":
                    headList.extend(["HDR",line[headerRecord.index("FileDate")],
                                     "",line[headerRecord.index("FileName")]])
                    footList.extend(["FTR",line[headerRecord.index("FileDate")],""])
            
            # Add record counts to the header and footer
            headList[2] = recordCount    
            footList[2] = recordCount    
            
            # Write all information to text
            csvOut.writerow(headList)
            for line in dataList:
                csvOut.writerow(line)
            csvOut.writerow(footList)
            
            
if __name__ == "__main__":
        main(sys.argv[1], sys.argv[2])
