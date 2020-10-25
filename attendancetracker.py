import os, sys, re, csv, json, configparser
import pandas as pd
from os import walk
from fuzzywuzzy import process

config = configparser.ConfigParser()
try:
    config.read('attendancetracker.conf')
except IOerror:
    try:
        config.read('default.conf')
    except IOerror:
        print("You might want to create a configuration file.")

#Load sections from config
print('Please choose a section by typing a number between 1 and 9:')
for key in config['classlist']:
    print('({0}) '.format(key) + config['classlist'][key])
sectionselection = input()
section = config['classlist'][sectionselection]

directory = (config['localization']['root_dir'] + section + "/")

# Get directory list for "Zoom Reports" folder

month = "October"

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

# Reads in Aliases dictionary file
with open('aliases.json','r') as read_file:
        aliases = json.load(read_file)

# Uses fuzzy matching to pair usernames with names on the roster
# Asks for confirmation if the match ratio is less than 80
# TO DO: check aliases.json first +
# Save confirmed matches to aliases.json +
# Allow manual pairing of poor matches
# (e.g. "Samsung Galaxy Note 9")
def tracker(section,date):
    zoom = pd.read_csv(directory + month + "/" + date + " - " + section + ".csv")
    already_present = []


    # Records attendance to dataframe.
    # Sometimes, students may log out and log back in
    # Currently only checks if ONE of their sessions is >= 30
    # (not the sum of both sessions)
    def recorder(match,ind,prettyname,zoomname):
        time = zoom.at[ind, "Total Duration (Minutes)"]
        if time >= 30:
            attendance.at[match.index, date] = 'P'
            already_present.append(prettyname)
            print(prettyname, 'attendance recorded')
        elif prettyname not in already_present:
            attendance.at[match.index, date] = int(time)
            print(prettyname, 'only present for {0} minutes'.format(time))

    for index, i in enumerate(zoom["Name (Original Name)"]):
        fmatch = process.extractOne(i,attendance['Name'])
        if i in aliases:
            prettymatch = aliases.get(i)
            match = attendance[attendance['Name'].str.contains(prettymatch.strip(), na=False, regex=False)]['Name']
            recorder(match,index,prettymatch,i)
        elif fmatch[1] > 80:
            match = attendance[attendance['Name'].str.contains(fmatch[0].strip(), na=False, regex=False)]['Name']
            print("matched",i,"to",fmatch[0])
            recorder(match,index,fmatch[0],i)
        elif fmatch[1] > 50:
            print("Would you like to associate",fmatch[0],"with",i,"?")
            confirm = input("Enter y for yes or n for no\n")
            if confirm == "y":
                match = attendance[attendance['Name'].str.contains(fmatch[0].strip(), na=False, regex=False)]['Name']
                recorder(match,index,fmatch[0],i)
                print("matched",i,"to",fmatch[0])
                roster_name = fmatch[0].strip()
                aliases[i] = roster_name
        else:
            print("************************")
            print("No match detected for", i)
            print("************************")
        print()
    with open('aliases.json', 'w') as f:
        json.dump(aliases, f)
    print()

    # Writes to attendance sheet in xlsx format
    writer = pd.ExcelWriter(directory + section + ".xlsx", engine='xlsxwriter')
    attendance.to_excel(writer, sheet_name='Sheet1',index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    absent = workbook.add_format({'bg_color': '#FFC7CE',
                               'bold': True})
    visiting = workbook.add_format({'bg_color': '#ADD8E6',
                               'bold': True})
    worksheet.set_column('A:A',32)
    worksheet.conditional_format('B2:S34',{'type': 'cell',
                                             'criteria': '=',
                                             'value': '"A"',
                                             'format': absent})
    worksheet.conditional_format('B2:S34',{'type': 'cell',
                                            'criteria': 'between',
                                            'minimum': 0.1,
                                            'maximum': 31,
                                            'format': visiting})
    writer.save()

# Run attendance tracker for all unrecorded Zoom reports in the month folder
for date in sessions_recorded:
    if date not in attendance.columns:
        print(date)
        attendance[date] = ""
        tracker(section,date)
