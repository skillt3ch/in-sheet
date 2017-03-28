#! /usr/bin/env python

import imaplib
import email
import getpass
import re
import weasyprint
import sys
import getopt

reload(sys)
sys.setdefaultencoding('utf-8')

"""
If the email contains multiple payloads (plaintext/html), you must parse each
message separately to get the text content (body)
"""


def get_first_text_block(email_message_instance):
    maintype = email_message_instance.get_content_maintype()
    if maintype == 'multipart':
        for part in email_message_instance.get_payload():
            if part.get_content_maintype() == 'text':
                return part.get_payload(decode=True)
    elif maintype == 'text':
        return email_message_instance.get_payload(decode=True)


"""
Using the body of the current email, scrape the search result
by traversing specific HTML tags which are present in the current format.
"""


def get_val(source, search):
     i = source.find(search)
     new_msg = source[i:]

     start = new_msg.find('</b>') + 5
     end = new_msg.find("</td>") - 1

     return unicode(new_msg[start:end].strip(), errors='replace')


"""
Using the mail object and the current email's UID,
scrape all the data and return a new object
"""


def scrape_data(mail, uid):
    result, data = mail.uid('fetch', uid, '(RFC822)')
    raw_email = data[0][1]

    scraped = {}

    email_message = email.message_from_string(raw_email)

    msg = get_first_text_block(email_message)
    subject = email_message['Subject']

    print '*' * 40, '\n', subject

    i = msg.find("No:")
    scraped['REQ'] = msg[i + 4:i + 10]

    print 'Web Request No: %s' % scraped['REQ']

    scraped['LOCATION'] = get_val(msg, "Location:")
    scraped['ORG'] = get_val(msg, "Organisation:")
    scraped['CONTACT'] = get_val(msg, "Contact:")
    scraped['ADDRESS'] = get_val(msg, "Address:")
    scraped['EMAIL'] = get_val(msg, "Email:")
    scraped['CONTACT_NUM'] = get_val(msg, "Contact Number:")
    scraped['MODEL'] = get_val(msg, "Model:")
    scraped['SERIAL'] = get_val(msg, "Serial:")
    scraped['EXTRA'] = get_val(msg, "Accessories:")
    scraped['FAULT'] = get_val(msg, "Repair Required:")
    scraped['DAMAGE'] = get_val(msg, "Damage:")
    scraped['PREV_REPAIR'] = get_val(msg, "Repaired:")
    scraped['BACKUP'] = get_val(msg, "Backup:")
    scraped['FMI'] = get_val(msg, "Find My iPhone:")
    scraped['REPAIR_TYPE'] = get_val(msg, "Repair Type:")
    scraped['RETURN_GOODS'] = get_val(msg, "Return of Goods:")
    scraped['AGREED'] = get_val(msg, "Service Terms:")
    scraped['REQ_DATE_TIME'] = get_val(msg, "Time of Request:")

    saved_data = ''
    # saved_data += '<p><strong>Service Location</strong>: %s</p>' % LOCATION
    saved_data += '<p><strong>Organisation</strong>: %s</p>' % scraped['ORG']
    saved_data += '<p><strong>Contact</strong>: %s</p>' % scraped['CONTACT']
    saved_data += '<p><strong>Address</strong>: %s</p>' % scraped['ADDRESS']
    saved_data += '<p><strong>Email</strong>: %s</p>' % scraped['EMAIL']
    saved_data += '<p><strong>Contact Number</strong>: %s</p>' % scraped['CONTACT_NUM']
    saved_data += '<p><strong>Brand/Model</strong>: %s</p>' % scraped['MODEL']
    saved_data += '<p><strong>Serial Number</strong>: %s</p>' % scraped['SERIAL']
    saved_data += '<p><strong>Extra Accessories</strong>: %s</p>' % scraped['EXTRA']
    saved_data += '<p><strong>Fault/Repair Required</strong>: %s</p>' % scraped['FAULT']
    saved_data += '<p><strong>Visible Damage</strong>: %s</p>' % scraped['DAMAGE']
    saved_data += '<p><strong>Previously Repaired</strong>: %s</p>' % scraped['PREV_REPAIR']
    saved_data += '<p><strong>Data Backup</strong>: %s</p>' % scraped['BACKUP']
    saved_data += '<p><strong>Find My iPhone</strong>: %s</p>' % scraped['FMI']
    saved_data += '<p><strong>Service Repair Type</strong>: %s</p>' % scraped['REPAIR_TYPE']
    saved_data += '<p><strong>Return of Goods</strong>: %s</p>' % scraped['RETURN_GOODS']
    saved_data += '<p><strong>Agreed to Service Terms</strong>: %s</p>' % scraped['AGREED']
    saved_data += '<p><strong>Date/Time of Request</strong>: %s</p>' % scraped['REQ_DATE_TIME']

    saved_data.replace("\\r\\n", "").replace("\\", "")

    scraped['saved_data'] = saved_data

    return scraped


def main(argv):

    # Initialise variables
    SERVER = "outlook.office365.com"
    USER = "jonathan.vonkelaita@compnow.com.au/servicevic@compnow.com.au"
    TEMPLATE = "template.html"
    JOB_NO = 0
    FOLDER = "s"
    HELP = """insheet.py [-u <username>] [-j <job number>] [-f <folder>]
folder options: s|sd|i
\ts:\t'Service Requests'
\tsd:\t'Service Requests DONE'
\ti:\t'INBOX'
"""

    # Parse the arguments
    try:
        opts, args = getopt.getopt(argv, "u:j:f:h")
    except getopt.GetoptError:
        print HELP
        sys.exit()

    # Get all the options/flags
    for opt, arg in opts:
        if opt == "-h":
            print HELP
            sys.exit()
        elif opt == "-j":
            JOB_NO = arg
        elif opt == "-u":
            USER = arg + "@compnow.com.au/servicevic@compnow.com.au"
        elif opt == "-f":
            FOLDER = arg

    if FOLDER == "s":
        FOLDER = "Service Requests"
    elif FOLDER == "sd":
        FOLDER = "Service Requests DONE"
    elif FOLDER == "i":
        FOLDER = "INBOX"
    else:
        print "Error: mailbox not recognised."
        sys.exit()

    print "Job Number:", JOB_NO
    print "Mailbox:", FOLDER
    print "Username:", USER
    PASS = getpass.getpass()

    try:
        # Log into mail server with username and password
        mail = imaplib.IMAP4_SSL(SERVER)
        mail.login(USER, PASS)

        # Connect to FOLDER
        status, count = mail.select(FOLDER)

        # If we didn't pass in a specific job number, just go through all emails which
        # contain the Subject: Hardware Service Booking
        if JOB_NO == 0:
            result, data = mail.uid('search', None, '(HEADER Subject "Hardware Service Booking")')
        # If we passed in a specific job number, then search the emails only for that email.
        else:
            result, data = mail.uid('search', None, '(BODY "Request No: %s")' % JOB_NO)

        # Create array/list of emails which match our search query.
        id_list = data[0].split()
    except Exception as e:
        print 'Error: %s' % e
        sys.exit()

    # Iterate through list of emails
    for uid in id_list:
        try:
            # Scrape current email and save data in new object
            scraped = scrape_data(mail, uid)

            # Open the HTML template and store it as a string
            with open(TEMPLATE, 'r') as htmlFile:
                html = htmlFile.read()

            # Replace all occurrences of special placeholder strings with our scraped data
            html = html.replace("$REQ$", scraped['REQ']).replace("$CONTACT$", scraped['CONTACT']).replace("$ORG$", scraped['ORG']).replace("$DATA$", scraped['saved_data'])

            # Convert to PDF document
            doc = weasyprint.HTML(string=html)
            pdf = doc.write_pdf()

            with open('%s-in.pdf' % scraped['REQ'], 'w') as pdfFile:
                pdfFile.write(pdf)

            print '\nSuccessfully saved.\n\t>> %s-in.pdf\n' % scraped['REQ'], '*'*40, '\n'

        except Exception as e:
            print 'Error: %s' % e
            sys.exit()

main(sys.argv[1:])
