################################################################################
##
##  Program:   Directory watcher and Emailer for Windows
##  Developer: Brandon Helms
##  Version:   1.0
##
################################################################################

import os
import win32file
import win32con
import smtplib
from time import sleep, strftime
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

#SETUP
###############################################################################
username = '<email_acct_to_send_from>'	#Real account made for them
password = '<email_passwd>'		#Real password made for them
server = '74.125.21.108:587'		#Google's SMTP IP
ADMIN_EMAIL = ''			#Email account of ADMIN
OTHER_EMAIL = ''			#Our email account for tracking if needed

WATCHED_PATH = 'c:\\'			#Directory to monitor
LOGFILE = r'.\logfile.txt'              #Log path locally
EXT_TO_MONITOR = ['asp', 'php', 'html'] #Extensions to monitor for
###############################################################################

ACTIONS = {
	1 : "Created",
	2 : "Deleted",
	3 : "Updated",
	4 : "RenamedFrom",
	5 : "RenamedTo"
}

def windows_watch_path(watched_path):
	'''watcher function'''

	FILE_LIST_DIRECTORY = 0x0001
	try:
		hDir = win32file.CreateFile (
			watched_path
			, FILE_LIST_DIRECTORY
			, win32con.FILE_SHARE_READ | 
			  win32con.FILE_SHARE_WRITE | 
			  win32con.FILE_SHARE_DELETE
			, None
			, win32con.OPEN_EXISTING
			, win32con.FILE_FLAG_BACKUP_SEMANTICS
			, None
		)
	except:
		print [[watched_path, '', ACTIONS[2]]]
		return -1


	while(True):
		results = win32file.ReadDirectoryChangesW (
			hDir,
			1024, 
			True,
			win32con.FILE_NOTIFY_CHANGE_FILE_NAME |
			win32con.FILE_NOTIFY_CHANGE_DIR_NAME |
			win32con.FILE_NOTIFY_CHANGE_ATTRIBUTES |
			win32con.FILE_NOTIFY_CHANGE_SIZE |
			win32con.FILE_NOTIFY_CHANGE_LAST_WRITE |
			win32con.FILE_NOTIFY_CHANGE_SECURITY,
			None,
			None
		)

		files_changed = []
		sendAndLog(results)

def sendAndLog(result):
	'''This is used output to file and cmd'''
	for action, fn in result:
		var =  '%s %s:\t"%s"' %(strftime('%y-%m-%d %H:%M:%S'), ACTIONS[action], WATCHED_PATH + os.sep + fn)
		if fn[-3:] in EXT_TO_MONITOR:
			compose_email([ADMIN_EMAIL, OTHER_EMAIL,''], 'ALERT: Files have been modified', [[var,0]], '')
		f = open(LOGFILE, 'a')
		f.writelines(var + '\n')
		f.close()			
		print var	#For testing only
	

def send_email(smtp_address, usr, password, msg, mode):
	server = smtplib.SMTP(smtp_address)
	server.ehlo()
	server.starttls()
	server.ehlo()
	server.login(username,password)
    
	if (mode == 0 and msg['To'] != ''):
		server.sendmail(msg['From'],(msg['To']+msg['Cc']).split(","), msg.as_string())
	elif (mode == 1 and msg['Bcc'] != ''):
		server.sendmail(msg['From'],msg['Bcc'].split(","),msg.as_string())
	elif (mode != 0 and mode != 1):
		print 'error in send mail bcc'; print 'email cancelled'; exit()
	server.quit()

def compose_email(addresses, subject, body, files):
	to_address = addresses[0]
	cc_address = addresses[1]
	bcc_address = addresses[2]
	from_address = ''
	
	msg = MIMEMultipart()
	msg['Subject'] = subject
	msg['To'] = to_address
	msg['Cc'] = cc_address
	msg['From'] = from_address
	
	for text in body:
		if (text[1] == 0):
			text[1] = 'plain'
		elif (text[1] == 1):
			text[1] = 'html'
		else:
			print 'error in text kind'; print 'email cancelled'; exit()
		
		part = MIMEText(text[0], text[1])
		msg.attach(part)

	if (files != ''):
		file_list = files.split(',')
		for afile in file_list:
			part = MIMEApplication(open(afile, "rb").read())
			part.add_header('Content-Disposition', 'attachment', filename=afile)
			msg.attach(part)
			
	send_email(server, username, password, msg, 0)

	if (bcc_address != ''):
		msg['Bcc'] = bcc_address
		send_email(server, username, password, msg, 1)
		
	#print 'email sent'  #Used for testing

if __name__ == '__main__':
	windows_watch_path(WATCHED_PATH)
