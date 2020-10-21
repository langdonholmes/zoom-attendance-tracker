import os, sys, re, csv, json
import pandas as pd
from os import walk
from fuzzywuzzy import process

# Names of classes
sections = {
    '1': "LING 329-02",
    '2': "LING 329-04",
    '3': "LING 460-560",
    '4': "LING 486",
}

print('Please choose a section by typing a number between 1 and 4:')
for x in sections:
    print('({0}) '.format(x) + sections[x])
sectionselection = input()
section = sections[sectionselection]

directory = ("C:/Users/Langdon/CSULB/Sarvenaz Hatami - Fall 2020 - Class Attendance/" + section + "/")
test = ("C:/Users/Langdon/PycharmProjects/AttendanceTracker/venv/" + section)

# Get directory list for "Zoom Reports" folder

month = "September"

# For the selection section, check .xlsx for a column matching each date
# If no date found, call the attendance tracker function

# Reads Attendance spreadhseet as dataframe
attendance = pd.read_excel(directory + section + ".xlsx")

#Regex Matching date format
regexp = r"([A-Za-z]{3,4}[.]\s\d{1,2})"

# First, get list of columns. Extract dates from this list.
dates_recorded = []
for date in attendance.columns:
    m = re.match(regexp,date)
    if m:
        dates_recorded.append(m.group(1))

# Check for files in directory that match date format.
# Makea a list of zoom reports in that folder.

sessions_recorded = []
for (dirpath,dirnames,filenames) in walk(directory + month):
    for name in filenames:
        m = re.match(regexp,name)
        if m:
            sessions_recorded.append(m.group(1))

# Sorts list of dates with a cute function and prints them
def datesorter(date):
    day = re.sub(r'\D', "", date)
    return(int(day))
sessions_recorded.sort(key=datesorter)
print(sessions_recorded)
print()

# Records attendance to dataframe.

def recorder(match,ind,prettyname,zoomname):
    time = zoom.at[ind, "Total Duration (Minutes)"]
    if time >= 30:
        attendance.at[match.index, date] = 'P'
        already_present.append(prettyname)
        print(prettyname, 'attendance recorded')
    elif prettyname not in already_present:
        attendance.at[match.index, date] = str(time)
        print(prettyname, 'only present for {0} minutes'.format(time))

# Uses fuzzy matching to pair usernames with names on the roster
# Asks for confirmation if the match ratio is less than 80

def tracker(section,date):
    zoom = pd.read_csv(directory + month + "/" + date + " - " + section + ".csv")
    already_present = []

    for index, i in enumerate(zoom["Name (Original Name)"]):
        fmatch = process.extractOne(i,attendance['Name'])
        if fmatch[1] > 80:
            match = attendance[attendance['Name'].str.contains(fmatch[0].strip(), na=False, regex=False)]['Name']
            recorder(match,index,fmatch[0],i)
            print("matched",i,"to",fmatch[0])
        elif fmatch[1] > 50:
            print("Would you like to associate",fmatch[0],"with",i,"?")
            confirm = input("Enter y for yes or n for no\n")
            if confirm == "y":
                match = attendance[attendance['Name'].str.contains(fmatch[0].strip(), na=False, regex=False)]['Name']
                recorder(match,index,fmatch[0],i)
                print("matched",i,"to",fmatch[0])
        else:
            print("no match detected for", i)


    print()

    writer = pd.ExcelWriter(directory + section + ".xlsx")
    attendance.to_excel(writer, sheet_name='Sheet1',index=False)
    worksheet = writer.sheets['Sheet1']
    writer.save()

for date in sessions_recorded:
    if date not in attendance.columns:
        print(date)
        attendance[date] = ""
        tracker(section,date)
