#! /usr/bin/env python

import imaplib
import email
import getpass
import re
import pdfkit

SERVER = "outlook.office365.com"
USER = "jonathan.vonkelaita@compnow.com.au/servicevic@compnow.com.au"
print "Username: %s" % USER
PASS = getpass.getpass()

# note that if you want to get text content (body) and the email contains
# multiple payloads (plaintext/ html), you must parse each message separately.
# use something like the following: (taken from a stackoverflow post)
def get_first_text_block(email_message_instance):
    maintype = email_message_instance.get_content_maintype()
    if maintype == 'multipart':
        for part in email_message_instance.get_payload():
            if part.get_content_maintype() == 'text':
                return part.get_payload(decode=True)
    elif maintype == 'text':
        return email_message_instance.get_payload(decode=True)

def get_val(source, search):
	 i = source.find(search)
	 new_msg = source[i:]

	 start = new_msg.find('</b>') + 5
	 end = new_msg.find("</td>") - 1

	 return unicode(new_msg[start:end].strip(), errors='replace')


mail = imaplib.IMAP4_SSL(SERVER)
mail.login(USER, PASS)

status, count = mail.select("Service Requests")

# print mail.list()


result, data = mail.uid('search', None, '(HEADER Subject "Hardware Service Booking")')

id_list = data[0].split()
# latest_email_id = id_list[-1]

for uid in id_list:
	result, data = mail.uid('fetch', uid, '(RFC822)')
	raw_email = data[0][1]

	email_message = email.message_from_string(raw_email)

	msg = get_first_text_block(email_message)
	subject = email_message['Subject']

	print '*'*40, '\n', subject

	i = msg.find("No:")
	REQ = msg[i+4:i+10]

	print 'Web Request No: %s' % REQ

	LOCATION = get_val(msg, "Location:")
	ORG = get_val(msg, "Organisation:")
	CONTACT = get_val(msg, "Contact:")
	ADDRESS = get_val(msg, "Address:")
	EMAIL = get_val(msg, "Email:")
	CONTACT_NUM = get_val(msg, "Contact Number:")
	MODEL = get_val(msg, "Model:")
	SERIAL = get_val(msg, "Serial:")
	EXTRA = get_val(msg, "Accessories:")
	FAULT = get_val(msg, "Repair Required:")
	DAMAGE = get_val(msg, "Damage:")
	PREV_REPAIR = get_val(msg, "Repaired:")
	BACKUP = get_val(msg, "Backup:")
	FMI = get_val(msg, "Find My iPhone:")
	REPAIR_TYPE = get_val(msg, "Repair Type:")
	RETURN_GOODS = get_val(msg, "Return of Goods:")
	AGREED = get_val(msg, "Service Terms:")
	REQ_DATE_TIME = get_val(msg, "Time of Request:")

	html = ''

	saved_data = ''
	#saved_data += '<p><strong>Service Location</strong>: %s</p>' % LOCATION
	saved_data += '<p><strong>Organisation</strong>: %s</p>' % ORG
	saved_data += '<p><strong>Contact</strong>: %s</p>' % CONTACT
	saved_data += '<p><strong>Address</strong>: %s</p>' % ADDRESS
	saved_data += '<p><strong>Email</strong>: %s</p>' % EMAIL
	saved_data += '<p><strong>Contact Number</strong>: %s</p>' % CONTACT_NUM
	saved_data += '<p><strong>Brand/Model</strong>: %s</p>' % MODEL
	saved_data += '<p><strong>Serial Number</strong>: %s</p>' % SERIAL
	saved_data += '<p><strong>Extra Accessories</strong>: %s</p>' % EXTRA
	saved_data += '<p><strong>Fault/Repair Required</strong>: %s</p>' % FAULT
	saved_data += '<p><strong>Visible Damage</strong>: %s</p>' % DAMAGE
	saved_data += '<p><strong>Previously Repaired</strong>: %s</p>' % PREV_REPAIR
	saved_data += '<p><strong>Data Backup</strong>: %s</p>' % BACKUP
	saved_data += '<p><strong>Find My iPhone</strong>: %s</p>' % FMI
	saved_data += '<p><strong>Service Repair Type</strong>: %s</p>' % REPAIR_TYPE
	saved_data += '<p><strong>Return of Goods</strong>: %s</p>' % RETURN_GOODS
	saved_data += '<p><strong>Agreed to Service Terms</strong>: %s</p>' % AGREED
	saved_data += '<p><strong>Date/Time of Request</strong>: %s</p>' % REQ_DATE_TIME

	saved_data = saved_data.replace("\\r\\n", "")

	with open('printtest.html', 'r') as htmlFile:
		html = htmlFile.read()

	html=html.replace("$REQ$", REQ).replace("$CONTACT$", CONTACT).replace("$ORG$", ORG).replace("$DATA$", saved_data)

	options = {
    'page-size': 'A4',
    'margin-top': '1mm',
    'margin-right': '1mm',
    'margin-bottom': '0mm',
    'margin-left': '1mm',
    }

	pdfkit.from_string(html, '%s-in.pdf' % REQ, options=options)


	#with open('%s-in.html' % REQ, 'w') as insheet:
	#	insheet.write(html)



	

	print saved_data

	"""

	with open('%s-in.txt' % REQ, 'w') as file:
		file.write(saved_data)

	with open('Call - %s.html' % REQ, 'w') as file:
		file.write(msg)
	"""

	print '\nSuccessfully saved.\n\t>> %s-in.pdf\n' % REQ, '*'*40, '\n'
