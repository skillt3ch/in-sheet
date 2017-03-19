#! /usr/bin/env python

import imaplib
import email
import getpass
import re

SERVER = "outlook.office365.com"
USER = "email@company.com"
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

	 start = new_msg.find('">') + 3
	 end = new_msg.find("<o:p") - 1

	 return new_msg[start:end]


mail = imaplib.IMAP4_SSL(SERVER)
mail.login(USER, PASS)

status, count = mail.select("INBOX")


result, data = mail.uid('search', None, '(HEADER Subject "Hardware Service Booking")')

id_list = data[0].split()
latest_email_id = id_list[-1]

for uid in id_list:
	result, data = mail.uid('fetch', uid, '(RFC822)')
	raw_email = data[0][1]

	email_message = email.message_from_string(raw_email)

	msg = get_first_text_block(email_message)
	subject = email_message['Subject']

	print '*'*20, '\n', subject

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


	saved_data = ''
	saved_data += 'Service Location: %s\n' % LOCATION
	saved_data += 'Organisation: %s\n' % ORG
	saved_data += 'Contact: %s\n' % CONTACT
	saved_data += 'Address: %s\n' % ADDRESS
	saved_data += 'Email: %s\n' % EMAIL
	saved_data += 'Contact Number: %s\n' % CONTACT_NUM
	saved_data += 'Brand/Model: %s\n' % MODEL
	saved_data += 'Serial Number: %s\n' % SERIAL
	saved_data += 'Extra Accessories: %s\n' % EXTRA
	saved_data += 'Fault/Repair Required: %s\n' % FAULT
	saved_data += 'Visible Damage: %s\n' % DAMAGE
	saved_data += 'Previously Repaired: %s\n' % PREV_REPAIR
	saved_data += 'Data Backup: %s\n' % BACKUP
	saved_data += 'Find My iPhone: %s\n' % FMI
	saved_data += 'Service Repair Type: %s\n' % REPAIR_TYPE
	saved_data += 'Return of Goods: %s\n' % RETURN_GOODS
	saved_data += 'Agreed to Service Terms: %s\n' % AGREED
	saved_data += 'Date/Time of Request: %s' % REQ_DATE_TIME

	saved_data = saved_data.replace("\\r\\n", "")

	print saved_data

	with open('%s-in.txt' % REQ, 'w') as file:
		file.write(saved_data)

	with open('Call - %s.html' % REQ, 'w') as file:
		file.write(msg)
	
	print '\nSuccessfully saved.\n\t>> Call - %s.html\n\t>> %s-in.txt\n' % (REQ, REQ), '*'*20, '\n'
