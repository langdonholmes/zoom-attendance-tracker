import os, sys, re, csv
import pandas as pd

# Names of classes
sections = {
    '1': "LING 329-02",
    '2': "LING 329-04",
    '3': "LING 460-560",
    '4': "LING 486",
}

# for file in os.listdir("Zoom Reports"):
#     date, therest = file.split(' - ')
#     section = therest.split('.csv', 1)[0]
#     print(date)
#     print(section)

print('Please choose a section by typing a number between 1 and 4:')
for x in sections:
    print('({0}) '.format(x) + sections[x])
sectionselection = input()

# Get directory list for "Zoom Reports" folder
# Extract dates from filenames in format MONTH. # where MONTH is a four-character abbreviation.

dates_recorded = []
date = "Choose a Date"

# For the selection section, check .xlsx for a column matching each date
# If no date found, call the attendance tracker function

section = sections[sectionselection]

# Reads files located in venv folder to Pandas as dataframes
directory = ("Choose a directory" + section)
attendance = pd.read_excel(directory + "/" + section + ".xlsx")

#Regex Matching date format
regexp = r"([A-Za-z]{3,4}[.]\s\d{1,2})"

# First, get list of columns. Extract dates from this list.
for name in attendance.columns:
    m = re.match(regexp,name)
    if m:
        dates_recorded.append(m.group(1))

# Check for files in directory that match date format. If some date does not match any in the list,
# make this string the variable "date"
from os import walk

regexp = r"([A-Za-z]{3,4}[.]\s\d{1,2})"
files = []
for (dirpath,dirnames,filenames) in walk(directory):
    for name in filenames:
        files.append(name)

dates_attended = []
for file in files:
    m = re.match(regexp,file)
    if m:
        dates_attended.append(m.group(1))

def tracker(section,date):
    zoom = pd.read_csv(directory + "/" + "September/" + date + " - " + section + ".csv")
    # Create new column named date if it does not already exist
    if date not in attendance.columns:
        attendance[date] = ""



# For each participant in the zoom meeting, checks for a match in the roster (by looking at aliases)
# Stores either the real name in the roster or an empty set as match
# If match contains a string, ouputs "P" under student's name/date to attendance dataframe
# If match is empty set (no match found), outputs screenname for manual verification.

#Initialize lists for any unidentified screennames or screennames present for < 30 minutes
    undercoveragents = []
    tourists = {}

    for index, i in enumerate(zoom["Name (Original Name)"]):
        x = 1
        match = attendance[attendance['Alias 1'].str.contains(i, na=False, regex=False)]['Name']
        # If a match is not found, this will check a list of alternative Aliases until a KeyError is raised (signalling
        # that nobody has x number of Aliases recorded / no such column "Alias X" exists. If this occurs,
        # break the while loop and append this screenname to list object undercoveragents.
        # Undercover agents are manually paired to a name in the roster and recorded
        # under Aliases (in the attendance .xlsx),
        # so no manual intervention is never needed more than once for a given screenname.
        # Would be fun to use RegEx at least for clearcut variations such as (name + XYZ) or (NAME) or (NA ME ) etc.
        while match.empty == True:
            x += 1
            try:
                match = attendance[attendance['Alias ' + str(x)].str.contains(i, na=False, regex=False)]['Name']
            except KeyError:
                break
        if match.empty == True:
            undercoveragents.append(i)
        elif zoom.at[index, "Total Duration (Minutes)"] >= 30:
            attendance.at[match.index, date] = 'P'
            print(match.to_string(), 'attendance recorded')
        else:
            print(match.to_string(), 'only present for {0} minutes'.format(zoom.at[index, "Total Duration (Minutes)"]))
            tourists.update({i:zoom.at[index, "Total Duration (Minutes)"]})

    print("\nAttendance not recorded for:\n")
    for covertoperative in undercoveragents:
        print(covertoperative)

    if len(tourists) > 0:
        print("\nAttended for less than 30 minutes:\n")
    for i in tourists:
        print(i, tourists[i], 'minutes')

    writer = pd.ExcelWriter("C:/Users/Langdon/CSULB/Sarvenaz Hatami - Fall 2020 - Class Attendance/" + section + "/" + section + ".xlsx")
    attendance.to_excel(writer, sheet_name='Sheet1',index=False)
    worksheet = writer.sheets['Sheet1']
    writer.save()

tracker(section,date)
