#Flask required imports
from flask import Flask
from flask import render_template
from flask import request
import io
import os

#ics file related imports
from icalendar import Calendar, Event
import pytz
from datetime import datetime
import tempfile

#Send Email related imports
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#General data getter/receiver/handler/displayer imports (numpy, pandas, excel, HTML display, etc.)
import numpy as np
import pandas as pd
import xlrd
import xlwt
import xlsxwriter
import openpyxl

#Activate json and csv for further use in the assignment
import json
import csv

app = Flask(__name__)

#Hyperlink Creation, for welcome_page link redirectory
@app.route('/')
def index():
	return render_template('welcome_page.html',name=None)

@app.route('/curr_meeting')
def curr_meeting():
    return render_template('curr_meeting.html',name=None)

@app.route('/new_meeting')
def new_meeting():
    return render_template('new_meeting.html',name=None)

@app.route('/view_personal')
def view_personal():
    return render_template('view_personal.html',name=None)

@app.route('/modify_personal')
def modify_personal():
    return render_template('modify_personal.html',name=None)

@app.route('/new_attendee')
def new_attendee():
    return render_template('new_attendee.html',name=None)

@app.route('/create_ics')
def create_ics():
    return render_template('create_ics.html',name=None)

@app.route('/send_email_page')
def send_email_page():
    return render_template('send_email.html',name=None)


#The followings are URLs which actually do the data transfer/creation/modification
@app.route('/curr_meeting_show', methods=['GET'])
def curr_meeting_show():
    # Show Current Meeting Schedule
    
    currMeetings = pd.read_csv('meeting_info.csv')
    return render_template('curr_meeting.html',currMeetings = currMeetings.to_html())

@app.route('/new_meeting_successful',methods=['POST'])
def new_meeting_successful():
    #Add New Meeting to meeting_info.csv
    
    #--Begin New Meeting Data Processing
    eventName = request.form.get('name')
    location = request.form.get('location')
    startY = request.form.get('startY')
    startM = request.form.get('startM')
    startD = request.form.get('startD')
    startH = request.form.get('startH')
    startMM = request.form.get('startMM')
    description = request.form.get('description')
    company = request.form.get('company')
    attendees = request.form.get('attendees')
    startTime = startY+"-"+startM+"-"+startD+"    "+startH+":"+startMM
    attendees.replace(" ","")
    #--End Data Processing
    
    #Read current meeting info
    currMeeting = pd.read_csv('meeting_info.csv')
    currMeeting = currMeeting[["Meeting Name","Location","Start Time","Company/Speaker","Description","Attendee Emails"]]
    
    #Add new meeting to the group
    newmeeting = [eventName,location,startTime,company,description,attendees]
    currMeeting.loc[-1] = newmeeting
    currMeeting.index = currMeeting.index + 1 
    currMeeting = currMeeting.sort_index()
    currMeeting.to_csv('meeting_info.csv')
    
    return render_template('meeting_successful.html', name=None)


@app.route('/view_personal_successful',methods=['POST','GET'])
def view_personal_successful():
    email = str(request.form.get('email'))
    pin = int(request.form.get('pin'))
    
    #--Begin Verify Email and Pin
    currPeople = pd.read_csv('people_info.csv')
    email_exists = currPeople["Email"].str.contains(email)
    flag = False
    for i in range(len(email_exists)):
        if (email_exists[i]):
            if (int(currPeople["Pin"][i]) == pin):
                flag = True
    
    error = "Cannot found such email & pin combination"
    if (flag == False):
        return json.dumps({ "error": error }), 200
    #--End Verify Email and Pin
    
    
    #Create Excel file for this specific user
    xlsName = "individual_calendar/"+email+".xlsx"
    currMeeting = pd.read_csv('meeting_info.csv')
    AllAttendee = currMeeting["Attendee Emails"]
    
    workbook = xlsxwriter.Workbook(xlsName)
    sheet = workbook.add_worksheet()
    sheet.activate()
        #if (os.path.isfile(xlsName) == False):
    sheet.write('A1','Meeting Name')
    sheet.write('B1','Location')
    sheet.write('C1','Start Time')

    j = 0
    for i in range(len(AllAttendee)):
        if (email in AllAttendee[i]):
            j += 1
            #Update user's current meeting schedule
            sheet.write(j,0,currMeeting["Meeting Name"][i])
            sheet.write(j,1,currMeeting["Location"][i])
            sheet.write(j,2,currMeeting["Start Time"][i])
                
    workbook.close()
    
    df = pd.read_excel(xlsName)
    csv_schedule = df.to_csv("individual_calendar/"+email+".csv", encoding='utf-8')
    personal_schedule = pd.read_csv("individual_calendar/"+email+".csv")
    
    return render_template('personal_schedule.html', email=email, personal_schedule = personal_schedule.to_html())


@app.route('/attendee_successful', methods = ['POST'])
def attendee_successful():
    #Add new attendee profile
    
    name = request.form.get('name')
    email = request.form.get('email')
    pin = request.form.get('pin')
    company = request.form.get('company')
    title = request.form.get('title')
    remarks = request.form.get('remarks')
    
    currPeople = pd.read_csv('people_info.csv')
    currPeople = currPeople[["Name","Email","Pin","Company","Title","Remarks"]]
    
    currPeople.loc[-1] = [name,email,pin,company,title,remarks]
    currPeople.index = currPeople.index + 1 
    currPeople = currPeople.sort_index()
    
    currPeople.sort_index().to_csv('people_info.csv')
    currPeople.sort_values('Name').to_csv('people_info_sorted.csv')

    return render_template('attendee_successful.html', name=None)


@app.route('/drop_meeting',methods=['POST'])
def drop_meeting():
    #Drop the meeting, identity verification by email and PIN
    
    email = str(request.form.get('email'))
    pin = int(request.form.get('pin'))
    
    currPeople = pd.read_csv('people_info.csv')
    email_exists = currPeople["Email"].str.contains(email)
    flag = False
    for i in range(len(email_exists)):
        if (email_exists[i]):
            if (int(currPeople["Pin"][i]) == pin):
                flag = True
    
    error = "Cannot found such email & pin combination"
    if (flag == False):
        return json.dumps({ "error": error }), 200
    
    xlsName = "individual_calendar/"+email+".xlsx"
    
    df = pd.read_excel(xlsName)
    csv_schedule = df.to_csv("individual_calendar/"+email+".csv", encoding='utf-8')
    personal_schedule = pd.read_csv("individual_calendar/"+email+".csv")
    
    return render_template('drop_meeting.html', personal_schedule = personal_schedule.to_html())


@app.route('/drop_successful',methods=['POST','GET'])
def drop_successful():
    #Drop the meeting by typing meeting name and email address
    
    name = request.form.get('meeting_name')
    email = request.form.get('email')
    
    currMeeting = pd.read_csv('meeting_info.csv')
    currMeeting_name = currMeeting["Meeting Name"]
    
    for i in range(len(currMeeting_name)):
        meetingName = currMeeting_name[i]
        if (meetingName.replace(" ","") == name.replace(" ","")):
            attendeeList = currMeeting["Attendee Emails"][i]
            attendeeList = attendeeList.replace(email,",")
            attendeeList = attendeeList.replace(",,",",")
            
            if (len(attendeeList) > 0):

                currMeeting["Attendee Emails"][i] = attendeeList
        
    currMeeting.to_csv("meeting_info.csv")
    
    return render_template('drop_meeting_successful.html', name=None)
    

@app.route('/create_events',methods = ['POST'])
def create_events():
    #For additional purposes, user can create an ics file themselves from scratch
    
    error = None
    if (request.method != 'POST'):
        return render_template('send_email.html', error = error)
    
    cal = Calendar()
    event = Event()
    
    organizer = "mailto:"+request.form.get('organizer_email')
    eventName = request.form.get('name')
    location = request.form.get('location')
    startY = int(request.form.get('startY'))
    startM = int(request.form.get('startM'))
    startD = int(request.form.get('startD'))
    startH = int(request.form.get('startH'))
    startMM = int(request.form.get('startMM'))
    startS = 0
    endY = int(request.form.get('endY'))
    endM = int(request.form.get('endM'))
    endD = int(request.form.get('endD'))
    endH = int(request.form.get('endH'))
    endMM = int(request.form.get('endMM'))
    endS = 0
    time_zone = pytz.timezone(request.form.get('Time_Zone'))
    description = request.form.get('description')
    
    event.add('summary', eventName)
    event.add('location', location)
    event.add('dtstart', datetime(startY,startM,startD,startH,startMM,startS,tzinfo = time_zone))
    event.add('dtend', datetime(endY,endM,endD,endH,endMM,endS,tzinfo = time_zone))
    event.add('dtstamp', datetime(startY,startM,startD,0,0,0,tzinfo=time_zone))
    event.add('description', description)
    #event.add('attendee', attendees)
    event.add('organizer', organizer)
    
    cal.add_component(event)
    
    #directory = tempfile.mkdtemp()
    directory = "/Users/eliye/AgendaProject"
    f = open(os.path.join(directory, request.form.get('file_name')), 'wb')
    f.write(cal.to_ical())
    f.close()
    
    return render_template('event_successful.html', name=None)


@app.route('/send_email', methods = ['POST'])
def send_email():
    #Send email to others with attachments of any kind
    error = None
    if (request.method != 'POST'):
        return render_template('send_email.html', error = error)
    
    fromaddr = request.form.get('fromaddr')
    toaddr = request.form.get('toaddr')
    cc = request.form.get('cc')
    pwd = request.form.get('password')
    
    msg = MIMEMultipart()
    
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Cc'] = cc
    msg['Subject'] = "An invitation by Python"
    
    body = request.form.get('body')
    
    msg.attach(MIMEText(body, 'plain'))
    
    filename = request.form.get('attachment')
    attachment = open("*/AgendaProject/"+filename, "rb")
    
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    msg.attach(part)

    server = smtplib.SMTP('smtp-mail.outlook.com', 587)
    server.starttls()
    server.login(fromaddr, pwd)
    text = msg.as_string()
    server.sendmail(fromaddr, [toaddr,cc], text)
    server.quit()
    
    return render_template('send_successful.html', name=None)
    
if __name__ == '__main__':
    app.run()
