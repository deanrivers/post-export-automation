# http://itsecmedia.com/blog/post/2016/python-send-outlook-email/

import win32com.client
from win32com.client import Dispatch, constants
import os
import datetime
from color_class import bcolors
import getpass


year = datetime.date.today().year
current_time = datetime.datetime.now().time().strftime('%H:%M:%S')

if current_time > '00:00' and current_time < '12:00':
    greeting = 'Good morning,'
elif current_time > '12:00':
    greeting = 'Good afternoon,'
else:
    greeting = 'Hello,'

def send_email(projectNumber):

    #const=win32com.client.constants
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = projectNumber+' '+"- Partial Data"
    # newMail.Body = "I AM\nTHE BODY MESSAGE!"
    newMail.BodyFormat = 2 # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
    path = os.path.join(r'P:\TAB\Tab Export',str(year),projectNumber)

    #determine who is using the computer
    current_user = getpass.getuser()

    #dictioanry of employees
    programming_employees = {
        "DRivers":{
            "name":"Dean Rivers",
            "title":"Online Project Supervisor",
            "phone": "201-840-5272",
            "CC":'Tom.Cariello@mvrg.com; Jose.Correa@mvrg.com;Brandon.Casasnovas@mvrg.com'
        },
        "BCasasnovas":{
            "name":"Brandon Casasnovas",
            "title":"Online Project Supervisor",
            "phone": "201-840-5266",
            "CC":'Tom.Cariello@mvrg.com;Jose.Correa@mvrg.com;Dean.Rivers@mvrg.com'
        },
        "TCariello":{
            "name":"Thomas Cariello",
            "title":"Director of Programming",
            "phone": "201-840-5258",
            "CC":'Jose.Correa@mvrg.com;Brandon.Casasnovas@mvrg.com;Dean.Rivers@mvrg.com'
        },
        "CChao":{
            "name":"Cheyenne Chao",
            "title":"Online Project Supervisor",
            "phone": "201-840-5274",
            "CC":'Jose.Correa@mvrg.com;Brandon.Casasnovas@mvrg.com;Dean.Rivers@mvrg.com;Tom.Cariello@mvrg.com'
        },
        "JCorrea":{
            "name":"Jose Correa",
            "title":"Head of Programming",
            "phone":"201-840-5276",
            "CC":'Brandon.Casasnovas@mvrg.com;Dean.Rivers@mvrg.com;Tom.Cariello@mvrg.com;'
        }
    }

    try:
        email_from = programming_employees[current_user]["name"]
        title = programming_employees[current_user]["title"]
        phone = programming_employees[current_user]["phone"]
        cc = programming_employees[current_user]["CC"]

    except:
        email_from = "Unknown User"
        title = "Unknown"
        phone = "Unknown"

    #image = '../assets/mvrg.gif'
    
    newMail.HTMLBody = ('<HTML><BODY>'+greeting+'<br><br> Partial data has been posted:'
    '<br><br>'+path+'<br><br>'
    'Thanks,'
    '<br><br>'
    '<b>'+email_from+'</b>'+
    '<br><img src="https://fm.myopinionnow.com/app/projects/mvrg.gif">'+
    '<br>'+
    title+
    '<br>'
    '115 River Road Suite 105'
    '<br>'
    'Edgewater, NJ 07020'
    '<br>'+
    phone +' | 201-840-6502'
    '</BODY></HTML>'
    )

    newMail.To = "John.Valente@mvrg.com; Valentine.Collis@mvrg.com; Richard.Mallet@mvrg.com"
    newMail.CC = cc
    #email_addresses = ["dean.rivers@mvrg.com","deanrivers2@mvrg.com"]

    #attachment1 = r"C:\Temp\example.pdf"
    #newMail.Attachments.Add(Source=attachment1)

    newMail.display()
    print('.')
    print('.')
    print('.')
    print(bcolors.OKBLUE+'The email has been prepped and is ready to send via Microsoft Outlook.'+bcolors.ENDC)
    #newMail.Send()