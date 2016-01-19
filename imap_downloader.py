#!/usr/bin/python
# -*- coding: utf-8 -*-
'''
Imap-Downloader: small script to download emails from an IMAP-Server and 
store them as text files, attachments are stored in their original format.
'''

import imaplib
import sys
import email
import email.header
import time
import datetime
import re
import mimetypes
import unicodedata
import os




# configuration
directory_for_year = True
directory_for_month = True	# month of year: 1..12



def sanatize(s):
	'''
	Simple conversion to ASCII-characters, e.g. 'Ä' to 'A'
	Return only characters that are alphanumeric or in {-, _, .} and discard others.
	'''
	s = unicodedata.normalize('NFKD',unicode(s,'utf-8', 'replace')).encode('ASCII','ignore')
	s = s.strip('-_ .')
	s = s.replace(' ','-')
	sanatized = ''.join((x for x  in s if x.isalnum() or x in '-_.'))
	return sanatized

def write_to_file(directory, filename, data):
	'''
	Create dir if nonexistant, sanatize filename, write file as binary or textfile (if .txt,.html,.htm)
	'''
	if data:
		filename = sanatize(filename)
		if len(filename) > 143:
			filename = filename[:123] + "..." + filename[-15:]
			#print ("new filenamelength: " + str(len(filename)))
		if not os.path.exists(directory): os.makedirs(directory)
		mode = 'w' if filename.endswith(('.txt','.htm','html')) else 'wb'
		f = open(os.path.join(directory,filename),mode)
		f.write(data)
		f.close()


def connectImap(host, username, password, port = 993):
	'''
	Connect to IMAP4_SLL, return imap
	'''
	M = imaplib.IMAP4_SSL(host)
	M.login(username,password)
	return M


def getMessageUids(imap4, directory):
	'''
	Get and return all UIDS in the passed directory
	'''
	for folder in M.list()[1]:
		print "\t" + str(folder).split(" ")[-1]
	selected_folder = raw_input("Please choose one of the above folders: ")
	inboxcount = M.select(selected_folder)[1][0]
	status, uids = M.search(None,'ALL')
	uids = uids[0].split()
	return uids


def downloadMessages(iinmap4, uids, process_message):
	"""
	Download each message (idenentified by its uid) seperately and process it 
	afterwards.
	"""
	total_amount = str(len(uids))
	for i in uids:
		print('Fetching message No.' + str(i)+'/' + total_amount + '...')
		mail = M.fetch(str(i),'(RFC822)')[1][0][1]
		process_message(mail)



def process_message(mail):
	"""
	Parses the content of each message

	@param mails 	iterable of email.message
	"""
	message = email.message_from_string(mail)	#parsing metadata
	datetuple = email.utils.parsedate_tz(message.__getitem__('Date'))
	filedirectory = basedirectory
	if not datetuple:
		datetuple = email.utils.parsedate_tz(message.__getitem__('Delivery-date'))
	if directory_for_year: 
		filedirectory = os.path.join(filedirectory, str(datetuple[0]))
	if directory_for_month:
		filedirectory = os.path.join(filedirectory, str(datetuple[1])) 
	dateposix = email.utils.mktime_tz(datetuple)
	localdate = datetime.datetime.fromtimestamp(dateposix)
	datestring = localdate.strftime('%Y%m%d-%H%M') # +'-'+'-'.join(time.tzname) #
	sender = email.utils.parseaddr(message['To'])[1].replace('@','_').replace('.','-')
	subject = email.header.decode_header(message['Subject'])[0][0]
	filename = datestring + '_' + sender[:60] + '_' + subject[:60]

	# parsing mail content
	mailstring = ''
	for headername, headervalue in message.items():
		mailstring += headername + ': ' + headervalue + '\r\n'	# add \r\n or
	if message.get_content_maintype() == 'text':
		mailstring += message.get_payload(decode=True)

	# handle multipart: 
	elif message.get_content_maintype() == 'multipart':
		partcounter = 0
		for part in message.walk():
			if part.get_content_maintype() == 'text':	# also: text/html
				for header, value in part.items():
					mailstring += header + ': ' + value + '\r\n'
					mailstring += '\r\n' + part.get_payload(decode=True) + '\r\n'
			# skip multipart containers
			elif part.get_content_maintype() != 'multipart':
				partcounter += 1
				try:
					attachmentname = email.header.decode_header(part.get_filename())[0][0]
				except:
					attachmentname = ""
					print("Error when parsing filename.")
				if not attachmentname:
					ext = mimetypes.guess_extension(part.get_content_type())
					if not ext:
						ext = '.bin'	# use generic if unknown extension
					attachmentname = 'attachment' + str(partcounter) + ext
				attfilename = filename + '_' + attachmentname
				write_to_file(filedirectory, attfilename, part.get_payload(decode=True))
	write_to_file(filedirectory, filename+'.txt', mailstring)


if __name__ == '__main__':
	if len(sys.argv) == 5:
		host = sys.argv[1]
		username = sys.argv[2]
		password = sys.argv[3]
		basedirectory = sys.argv[4]
	else:		
		host = raw_input("Please enter your imap server: ")
		username = raw_input("Please enter your username: ")
		password = raw_input("Please enter your account's password: ")
		basedirectory = raw_input("Please enter the folder to save the mails in: ")
	

	M = connectImap(host,username,password)
	#print(str(M.list()))
	uids = getMessageUids(M, 'Sent')
	#print (len(uids))
	downloadMessages(M, uids, process_message)
