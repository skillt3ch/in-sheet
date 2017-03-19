#! /usr/bin/env python3

import imaplib
import email
# import getpass
import re

SERVER = "outlook.office365.com"
USER = "jonathan.vonkelaita@compnow.com.au"
PASS = "bLueivY7361"

mail = imaplib.IMAP4_SSL(SERVER)
mail.login(USER, PASS)

# print(mail.list())

mail.select("INBOX")

# result, data = mail.uid('search', None, "ALL")
# latest_email_uid = data[0].split()[-1]
# result, data = mail.uid('fetch', latest_email_uid, '(RFC822)')

# raw_email = data[0][1]

# print(data, latest_email_uid)

# search for all emails containing the subject title "PO Numbers"
# returns a 1 item array containing the list of UIDs of the emails which match the search string
# e.g. [b'310 353 1093 3289']
result, data = mail.uid('search', None, '(HEADER Subject "PO Numbers")')

# decode binary data to string
data[0] = data[0].decode('utf-8')
# convert to array and reverse (to sort from newest to oldest)
data[0] = reversed(data[0].split())

# iterate through each email UID
for elem in data[0]:
	# fetch current email
	result, data = mail.uid('fetch', elem, '(RFC822)')
	# grab raw email and convert to utf-8 encoded string
	raw_email = data[0][1].decode('utf-8')
	# parse to email object
	email_message = email.message_from_string(raw_email)
	# store body of email
	msg = email_message.get_payload()

	with open('%s.html' % elem, 'w') as file:
		file.write(msg)
	
	# print out message UID of each email
	print("### Message UID: %s ###" % (elem))

	try:
		# search body for "VMEM" and store the index
		i = msg.find("VMEM")
		# search for "<" (which is the first character after the end of the PO number)
		j = msg[i:].find("<")
		# vmem = i + 19  (19 characters between VMEM and the start of the PO number) -- "&nbsp; &nbsp;" in between.
		vmem = msg[i+19:i+j]
		# print out the subject title and VMEM PO Number
		print('%s\nVMEM:\t%s\n' % (email_message['Subject'], vmem))
	except:
		print("...\n")
		# print(msg)