"""
Author: Srideep Kar
"""

import time
from tokenize import Name
import win32com.client as win32
from pandas import *
from datetime import datetime

def create_dict_from_excel():
    """
    convert the employee data from Excel sheet to dictionary
    """
    print("Converting Excel Doc into Dictionary....")

    emp_xls = ExcelFile('Confirmation_Tracker.xls')
    emp_dict = emp_xls.parse(emp_xls.sheet_names[0])
    emp_dict = emp_dict.to_dict()
    return emp_dict

def time_tracker(emp_dict):
    """
    Track time and call send mail method a day prior of Confirmation Due date or confirmation initiation date
    """
    print("Running Time Tracker....")

    # fetch current time
    now = datetime.now()
    date_time_now = now.strftime("%b %d %H:%M:%S %Y")
    
    # confirmation Initiation Date
    operation="Confirmation Initiation Date"
    CID_emp_list = extract_emp_info(emp_dict, date_time_now, operation)

    # Confirmation Due Date
    operation="Confirmation Due Date"
    CDD_emp_list = extract_emp_info(emp_dict, date_time_now, operation)

    return(CID_emp_list, CDD_emp_list)
    

def extract_emp_info(emp_dict, date_time_now, operation):
    """
    returns a list of number of those employees whose Confirmation Initiation Date or Confirmation Due Date is tomorrow
    """
    print("Extracting employee informations from database....")

    emp_list = []
    for emp_num in range (0,len(emp_dict[operation])):
        # fetch Date time
        date_time_emp = datetime.strftime(emp_dict[operation][emp_num], "%m/%d/%Y, %H:%M:%S")
        emp_time = datetime.strptime(date_time_emp, "%m/%d/%Y, %H:%M:%S")
        current_time = datetime.strptime(date_time_now, "%b %d %H:%M:%S %Y")       

        # calculate difference between both time
        time_remaining = emp_time - current_time
        # Add employee number to the emp_list if only 1 day remaining before Date
        if time_remaining.days == 0:
            emp_list.append(str(emp_num))
        
    return emp_list

def send_email(Subject, Content_Email, recipients):
    """
    Send Emails
    """
    print("Sending Mail process initiated....")

    try:
        outlook = win32.Dispatch('Outlook.Application')
        # Create new E-mail
        mail = outlook.CreateItem(0)

        # Subject of E-Mail
        mail.Subject = Subject

        # Content of Email
        mail.BodyFormat = 1

        # Mail recipients
        recipients = recipients[0] + ";" + recipients[1]
        mail.To = recipients

        # Attachments
        image = 'logo.jpg'

        # Add image into E-mail Body
        attachment = mail.Attachments.Add(image)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")

        # Create E-mail body
        mail.HTMLBody = Content_Email + "<html><body><img src=""cid:MyId1""></body></html>"

        # Send E-mail
        mail.Sensitivity  = 2
        #mail.Display()
        mail.Save()
        mail.Send()

        print("E-Mail Sent!!!")
    
    except Exception as e:
        print("Error!!!" + str(e))

def create_email():
    """
    Create Email remainder and send it one day prior of Confirmation Initiation Date & Confirmation Due Date to employee's respective managers & HRs
    """
    print("Writing E-Mail...")

    # convert the Excel file to a dict
    emp_dict = create_dict_from_excel()

    # track time
    CID_emp_list, CDD_emp_list = time_tracker(emp_dict)

    # Check if there any employee exist in CID_emp_list
    if len(CID_emp_list) > 0:
        operation="Confirmation Initiation Date"
        # Create Mail info and send it
        for emp_num in CID_emp_list:
            # create email body
            Content_Email = create_body(operation, emp_num, emp_dict)
            # create email subject
            Subject = create_subject(operation, emp_num, emp_dict)
            # email recipients i.e. managers & HR's email
            recipients = [emp_dict['HR Email'][int(emp_num)], emp_dict['Reporting Manager Email'][int(emp_num)]]
            # Send email
            send_email(Subject, Content_Email, recipients)
    
    # Check if there any employee exist in CDD_emp_list
    if len(CDD_emp_list) > 0:
        operation="Confirmation Due Date"
        # Create Mail info and send it
        for emp_num in CDD_emp_list:
            # create email body
            Content_Email = create_body(operation, emp_num, emp_dict)
            # create email subject
            Subject = create_subject(operation, emp_num, emp_dict)
            # Send email
            send_email(Subject, Content_Email, recipients)
    
def create_body(operation, emp_num, emp_dict):
    """
    Create Email body for given employee number
    """
    print("Filling Up E-Mail Body...")

    emp_name = emp_dict['Name'][int(emp_num)]

    emp_code = emp_dict['Emp Code'][int(emp_num)]

    msg = "Hi, This is an automated e-mail to notify you that, tomorrow is the " + operation + " of MR/MS/MRS " + emp_name + ". Employee ID: " + emp_code

    return msg

def create_subject(operation, emp_num, emp_dict):
    """
    Create subject of the email
    """
    print("Fulling Up E-Mail Subject")

    emp_name = emp_dict['Name'][int(emp_num)]

    msg = operation + " of " + emp_name

    return msg       

def email_loop():
    """
    Email loop which execute 24x7 and trigger the process at 1655799840th second of the day
    """
    print("Welcome to the Automated E-mail system ")
    print("Email System triggered...")
    while(True):
        # trigger the mail at 1655799840th second of the day
        if int(time.time()) == 1655799840:
            time.sleep(5)
            create_email()

if __name__=="__main__":
    try:
        email_loop()
    except KeyboardInterrupt:
        print("Automated Mail system Closed.")