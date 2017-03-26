#! /usr/bin/env python

import imaplib
import email
import getpass
import re
import weasyprint
import sys
import getopt

# encode everything to utf-8
reload(sys)
sys.setdefaultencoding('utf-8')

"""
If the email contains multiple payloads (plaintext/html), you must
parse each message separately to get text content (body):
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
Takes the email message (source) and a (search) term:
Searches the current email message for the value which is nested between some HTML tags
</b> and </td>. This current layout is specific to the in-sheet emails we receive.
"""
def get_val(source, search):
	 i = source.find(search)
	 new_msg = source[i:]

	 start = new_msg.find('</b>') + 5
	 end = new_msg.find("</td>") - 1

	 # strip all whitespace at the start and end of the string and
	 # convert it to unicode before returning
	 return unicode(new_msg[start:end].strip(), errors='replace')

"""
This is where the bulk of the scraping gets done.
It takes the mail object and the UID of the current email.
Using the UID, it grabs the raw data of the email and converts it to
an 'email' object. The body of the email is then stored in 'msg'.
All of the customer data / description / info is scraped and stored
in a new object (scraped), and converted to a newly formatted HTML
string. The HTML string is also saved to the scraped object.
The whole scraped object is returned, looking something like:
['LOCATION': 'VIC', 'ORG': 'Monash University', ..., 'saved_data': '<p>...</p>']
"""
def scrape_data(mail, uid):
	# fetch current email and store in "data"
	result, data = mail.uid('fetch', uid, '(RFC822)')
	# data is an array. Grab the raw email data from the array.
	raw_email = data[0][1]

	scraped = {}

	# create email_message object from the converted raw_email
	email_message = email.message_from_string(raw_email)

	# grab body and subject of email
	msg = get_first_text_block(email_message)
	subject = email_message['Subject']

	print '*'*40, '\n', subject

	# scrape all data and save in new object
	i = msg.find("No:")
	scraped['REQ'] = msg[i+4:i+10]

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

	# convert all scraped data into a newly formatted HTML string, and save to scraped object
	saved_data = ''
	#saved_data += '<p><strong>Service Location</strong>: %s</p>' % LOCATION
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

	# remove any \r\n and \ instances from the scraped data
	saved_data.replace("\\r\\n", "").replace("\\", "")

	scraped['saved_data'] = saved_data

	return scraped

def main(argv):

	# initialise values
	SERVER = "outlook.office365.com"
	USER = "jonathan.vonkelaita@compnow.com.au/servicevic@compnow.com.au"
	JOB_NO = 0
<<<<<<< HEAD
	FOLDER = "Service Requests"
=======
	TEMPLATE = "template.html"
>>>>>>> 26cebfab77cb79a9c438f152ee379a79b77f18e1

	# set up argument handler
	try:
		opts, args = getopt.getopt(argv, "u:j:f:h")
	except getopt.GetoptError:
		print "insheet.py [-u <username>] [-j <job number>]"

	for opt, arg in opts:
		if opt == "-h":
			print "insheet.py [-u <username>] [-j <job number>]"
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

<<<<<<< HEAD
	print "Job Number:", JOB_NO
	print "Mailbox:", FOLDER
=======
>>>>>>> 26cebfab77cb79a9c438f152ee379a79b77f18e1
	print "Username:", USER
	# get password
	PASS = getpass.getpass()

	# create instance of IMAP object and connect to the server
	mail = imaplib.IMAP4_SSL(SERVER)
	# using the username and password, we try to log in
	mail.login(USER, PASS)

	# Go to the Service Requests folder
	status, count = mail.select("Service Requests")

	# if we did not specify a job number, then grab all available jobs
	if JOB_NO == 0:
		result, data = mail.uid('search', None, '(HEADER Subject "Hardware Service Booking")')
	# otherwise, grab only the email whose body matches "Request No: " + job number
	else:
		result, data = mail.uid('search', None, '(BODY "Request No: %s")' % JOB_NO)

	# create a list of all email UIDs which match the above search.
	# Searching a single job number should return a list with only one element.
	id_list = data[0].split()

	# iterate through each UID in the list
	for uid in id_list:

		try:
			# call the scrape_data function, passing in our mail object and the UID of the current email.
			scraped = scrape_data(mail, uid)

			# store the HTML of our template to a string
			with open(TEMPLATE, 'r') as htmlFile:
				html = htmlFile.read()

			# using the template, replace all placeholders with the actual scraped data
			html = html.replace("$REQ$", scraped['REQ']).replace("$CONTACT$", scraped['CONTACT']).replace("$ORG$", scraped['ORG']).replace("$DATA$", scraped['saved_data'])

			# use weasyprint library to convert the HTML string into a pdf document
			doc = weasyprint.HTML(string=html)
			pdf = doc.write_pdf()

			# save the pdf document to the current directory using the "<jobnumber>-in.pdf" as the filename
			with open('%s-in.pdf' % scraped['REQ'], 'w') as pdfFile:
				pdfFile.write(pdf)

			# print out the success message
			print '\nSuccessfully saved.\n\t>> %s-in.pdf\n' % scraped['REQ'], '*'*40, '\n'

		# if there is any error executing the above code, it will print out the error it encountered
		except Exception as e:
			print 'Error: %s' % e

# call the main function passing in all arguments (not including the filename of the python script)
main(sys.argv[1:])