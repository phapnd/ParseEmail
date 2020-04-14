import imapclient
import pyzmail
import pprint
import os
from pyzmail import PyzMessage
import email

def connect(server, email, password):
	"""
	Get Imap account cconnection with server
	"""
	imapObj = imapclient.IMAPClient(server, use_uid= True, ssl=True)
	imapObj.login(email, password)
	imapObj.select_folder('INBOX', readonly=True)
	return imapObj

def print_tree(account):
    """
    Print folder tree
    """
    print(account.list_folders())

def get_unread(account):
    """
    Get unread emails
    """
    UIDs = account.search(['UNSEEN'])
    return UIDs

def get_pyzmail(account,id_email):
    """
    Get pyzmail email
    """
    raw_msg = account.fetch([id_email], [b'BODY[]', 'FLAGS'])
    msg =  pyzmail.PyzMessage.factory(raw_msg[id_email][b'BODY[]'])
    return msg

def get_stmail(account,id_email):
	"""
	Get standard format email
	"""
	messageParts = account.fetch([id_email], [b'BODY[]', 'RFC822'])
	#emailBody = messageParts[0][1]
	print("\nGET_STMAIL")
	print(messageParts)
	mail = email.message_from_bytes(messageParts[id_email][b'RFC822'])
	return mail

def download_attachment(mail,folder_name):
	"""
	Download attachment included in email, save in folder_name
	"""
	for part in mail.walk():
		if part.get_content_maintype() == 'multipart':
			# print part.as_string()
			continue
		if part.get('Content-Disposition') is None:
			# print part.as_string()
			continue
		fileName = part.get_filename()
		if bool(fileName):
		    filePath = os.path.join(folder_name, 'attachments', fileName)
		    if not os.path.isfile(filePath) :
		        print(fileName)
		        #print(part.get_payload(decode=True))
		        fp = open(filePath, 'wb')
		        fp.write(part.get_payload(decode=True))
		        fp.close()

def main():
    # Connection details
    server = 'imap.gmail.com'
    email = input("Email: \n")
    password = input("Password: \n")

    account = connect(server, email, password)

    # Print all subject emails
    UIDs = get_unread(account)
    for uid in UIDs:
    	msg = get_pyzmail(account,uid)
    	print("\nFrom: " + ''.join(msg.get_address("from")))
    	print("\nTo: " + ''.join(msg.get_address("to")))
    	print("\nSubject: " + msg.get_subject())
    	body_msg_txt = lambda msg: "" if  msg.text_part == None else  msg.text_part.get_payload().decode(msg.text_part.charset)
    	print("\nBody: ")
    	print(body_msg_txt(msg))
    	body_msg_html = lambda msg: "" if msg.html_part == None else msg.html_part.get_payload().decode(msg.html_part.charset)
    	print("\n")
    	download_attachment(get_stmail(account,uid),r"C://Users//hanht//Downloads//")
    	#print(body_msg_html(msg))
    	# Download Attachment


    account.logout()

if __name__ == '__main__':
    main()
