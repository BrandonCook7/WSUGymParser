'''
Developed by Brandon Cook
Last Build Date: 2/25/2021
'''
import win32com.client
import os
from datetime import datetime, timedelta

#Modifier Variables for User
ageOfEmails = 7 #This number represents how many days do you want to look back for the emails
includePastDue = False #Set to True if you want past appointments added, but they need to be within the ageOfEmail number
remindMe = False #Set to true if you want reminders before your appointment
remindTime = 15 #Change to how many minutes you want to be reminded, ignore this if you set remindMe to false
deleteAfterAdded = True #Will delete the email after it has been added to your calender
noRepeats = True #Not added yet but if you have delete emails to true you shouldn't run into this repeats

allDates = []#Creates a list for storing all meta data
def extract(days):
    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace("MAPI")
    for account in mapi.Accounts:
	    print(account.DeliveryStore.DisplayName)
    inbox = mapi.GetDefaultFolder(6)
    messages = inbox.Items
    received_dt = datetime.now() - timedelta(days)
    received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
    messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
    for m in messages:
        if m.Subject == "UREC Purchase Information":
            text = m.body
            gymInfo = textParser(text)
            allDates.append(gymInfo)
            if(deleteAfterAdded):
                m.Delete()#Delete the email after getting the data
            
def textParser(text):

    startTime = datetime.now()

    end = text.rfind("------------------------------------------------")
    if end == -1:
        raise ValueError("ERROR: Issue finding correct spot to trim, contact the dev the email format may have changed, DEV NOTE: Trim1")
    text = text[:end]
    start = text.rfind("------------------------------------------------")
    if start == -1:
        raise ValueError("ERROR: Issue finding correct spot to trim, contact the dev the email format may have changed, DEV NOTE: Trim2")

    loc = text.find("Student Recreation Center")
    if(loc == -1):#If email does not have a Student Recreation Center reservation it checks for a chinook appointment
        loc = text.find("Chinook Recreation Center")
        if(loc == -1):
            raise ValueError("ERROR: Unknown location scheduled")
            #TODO Get a Stephenson Gym reservation email to build support for their reservation system

    parsedText = text[loc:]
    #Getting Time Information
    eol = parsedText.find("\r\n")
    calenderLocation = parsedText[:eol]

    parsedText = parsedText[eol+2:]#+2 because \r\n add up to 2 characters
    eol = parsedText.find("\r\n")
    calenderDesc = parsedText[:eol]

    #This part is slightly different because the whole line is not needed only the second part
    parsedText = parsedText[eol+2:]#Start of last line

    startSpot = parsedText.find(":")#Finds the first ':' which is one space before the date


    eol = parsedText.find("\r\n")
    timeUnFormated = parsedText[startSpot+2:eol]#+2 because of the ':' and whitespace
    startTime = timeParser(timeUnFormated)#Converts time to a supported datetime format

    gymList = [calenderLocation, calenderDesc, startTime]
    return gymList
def timeParser(timeData):
    date_object = datetime.strptime(timeData, '%m/%d/%Y %I:%M:%S %p')
    #print(timeData)
    print(date_object)

    return date_object

def addEvent(gymList):
    outlook = win32com.client.Dispatch("Outlook.Application")
    app = outlook.CreateItem(1)
    date = gymList[2]
    
    #Checks if appointment is older than current datetime
    if((date < date.now()) and includePastDue == False):
        print("OLD APPOINTMENT")
        return
    #Reorganizing Date Format
    outlookTime = date.strftime("%m-%d-%y %H:%M")
    #TODO Check for Repeats
    #TODO Auto merge two appointments back to back
    app.Start = outlookTime
    app.Subject = gymList[1]
    app.Duration = 60
    app.Location = gymList[0]
    app.ReminderSet = remindMe
    app.ReminderMinutesBeforeStart = remindTime
    app.Save()
    return

extract(ageOfEmails)#Gets email data

for i in range(len(allDates)):#Uses email data to create appointments
    addEvent(allDates[i])
print("Finished Adding Appointments")