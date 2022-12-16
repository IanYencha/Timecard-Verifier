from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl import Workbook
import os
import pandas as pd
from datetime import timedelta, date, time, datetime
os.chdir("C:\\Users\\Ian\\OneDrive - University of Pittsburgh\\Career\\SCS Employment\\Timecard automation")

#TODO: Ask for file name, make output file the correct name (based on date)
#load login times and the schedule excel docs
wb1 = load_workbook("(RAW) 2022-12-04.xlsx",read_only=False)
report = wb1.active

keys = pd.read_excel(io="(KEYS) 2022-12-04.xlsx")
keys = keys.drop(columns=["idnt","agid","Duration", "Name"]) #remove useless cols
# keys = keys.sort_values(['User', 'Logon'])


# Traverse through report sheet, verifying schedule with keys sheet
report_row = 1

# Loops for each employee
username = "THIS IS A PLACEHOLDER"
while username != 'ity1': #TODO: add actual condition
    #get username and full name from curr row and col
    raw_cell_name = report.cell(row=report_row, column=1).value
    if (raw_cell_name == None):
        break
    allParts = raw_cell_name.split(sep="|")
    
    full_name = allParts[0][:-1]
    username = allParts[1][1:]
    print("Full name: " + full_name + "\nusername: " + username)
    
    user_logins = []
    user_logouts = []
    for row in range(len(keys.index)): #Get all logins by employee
        if keys.loc[row][0].casefold() == username.casefold() or keys.loc[row][0].casefold() == full_name.casefold():
            user_logins.append(keys.loc[row][1])
            user_logouts.append(keys.loc[row][2])
    user_logins.sort()
    user_logouts.sort()
    #print(user_logins + user_logouts)
    
    
    #Loops for each shift for an employee
    report_row += 1
    currVal = report.cell(row=report_row, column= 1).value
    while currVal != None: # None means end of this employees shifts, 
        #Look at single shift (or connected shifts)
        num_connected = 0
        date = report.cell(row=report_row, column=1).value
        start = datetime.combine(date, report.cell(row=report_row, column=2).value)
        end = datetime.combine(date, report.cell(row=report_row, column=3).value)
        # While end time is equal to start time of next shift and dates are equal, set end time to next shift end
        while (report.cell(row=report_row+1, column=1).value != None 
                and end >= datetime.combine(report.cell(row=report_row+1, column=1).value, report.cell(row=report_row+1, column=2).value) - timedelta(hours=1)): # Combines shifts within 1 hours of eachother
            report_row += 1
            num_connected += 1
            end = datetime.combine(date, report.cell(row=report_row, column=3).value)
        
        # check keys for earliest logon that matches
        lessThan = False
        closest_start_time = user_logins[0]
        for time in user_logins:
            time_diff = start - time
            if abs(time_diff) < abs(start - closest_start_time):
                closest_start_time = time
                
        startLate = closest_start_time - start > timedelta(minutes=6)

        # Check for closest logoff
        closest_end_time = user_logouts[0]
        for time in user_logouts:
            time_diff = end - time
            if abs(time_diff) < abs(end - closest_end_time):
                closest_end_time = time

        endEarly = end - closest_end_time > timedelta(minutes=6)
        print("Checking " + str(num_connected+1) + " shift(s): " + str(start) + " to " + str(end))
        #Edit sheet based on results
        if (closest_start_time > end or closest_end_time < start or 
            closest_end_time < end - timedelta(hours=2) or closest_start_time > start + timedelta(hours=2)): # FF0000 - red
            #print("No login for shift found")
            for row in range (report_row - num_connected, report_row + 1):
                for col in range(1,7):
                    report.cell(row=row, column=col).fill = PatternFill(fgColor="FF0000", fill_type="solid")
        else:
            if startLate: # FFFF00 - yellow
                #print("Employee started late: " + str(closest_start_time)) #TODO: implement if late for middle shift
                #Highlight Yellow
                for col in range(1,7): 
                    report.cell(row=report_row-num_connected, column=col).fill = PatternFill(fgColor="FFFF00", fill_type="solid")
                if num_connected > 0:
                    for row in range((report_row-num_connected) + 1, report_row):
                        for col in range(1,7):
                            report.cell(row=row, column=col).fill = PatternFill(fgColor="92D050", fill_type="solid")
                    if not endEarly:
                        for col in range(1,7):
                            report.cell(row=report_row, column=col).fill = PatternFill(fgColor="92D050", fill_type="solid")
                
                
                # Output time in col 7 for row
                zero_str = ""
                hour = closest_start_time.hour
                if (hour > 12):
                    hour -= 12
                if (closest_start_time.minute < 10):
                    zero_str = "0"
                start_time_str = str(hour) + ":" + zero_str + str(closest_start_time.minute)
                report.cell(row=report_row-num_connected,column=7).value = start_time_str
            if endEarly:
                #print("Employee left early: " + str(closest_end_time))
                currRow = report_row - num_connected
                while datetime.combine(report.cell(row=currRow, column=1).value, report.cell(row=currRow, column=3).value) < closest_end_time:
                    for col in range(1,7):
                            report.cell(row=currRow, column=col).fill = PatternFill(fgColor="92D050", fill_type="solid")
                    currRow += 1
                while currRow != report_row + 1:
                    for col in range(1,7):
                        report.cell(row=currRow, column=col).fill = PatternFill(fgColor="FFFF00", fill_type="solid")
                    currRow += 1

                # Output the time in col 8 for row
                zero_str = ""
                hour = closest_end_time.hour
                if (closest_end_time.minute < 10):
                    zero_str = "0"
                if hour > 12:
                    hour -= 12
                end_time_str = str(hour) + ":" + zero_str + str(closest_end_time.minute)
                report.cell(row=report_row, column=8).value = end_time_str
            if (not startLate and not endEarly): # 92D050 - Green
                #print ("Valid logins found for this shift")
                for row in range((report_row-num_connected), report_row+1):
                        for col in range(1,7):
                            report.cell(row=row, column=col).fill = PatternFill(fgColor="92D050", fill_type="solid")
        print()
        
        report_row += 1
        currVal = report.cell(row=report_row, column= 1).value
    
    report_row += 1

print("Processed all Shifts, Report completed")
Workbook.save(self=wb1, filename="Validated.xlsx")












 
#DEMO
# import openpyxl

# # open the Excel file
# wb = openpyxl.load_workbook("my_excel_file.xlsx")

# # select the active sheet
# sheet = wb.active

# # read the value of the cell at row 1, column 1
# cell_value = sheet.cell(row=1, column=1).value

# # change the value of the cell at row 1, column 1
# sheet.cell(row=1, column=1).value = "new value"

# # highlight the cell at row 1, column 1
# sheet.cell(row=1, column=1).style = "highlight"

# # save the changes to the Excel file
# wb.save("my_excel_file.xlsx")