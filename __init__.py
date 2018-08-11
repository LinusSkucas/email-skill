from mycroft import MycroftSkill, intent_file_handler
import sys
import imaplib
import email
import email.header

def list_new_email(account, folder, password, port, address):
    M = imaplib.IMAP4_SSL(str(address), port=int(port))
    M.login(str(account), str(password))
    M.select(str(folder))
    #TODO: PROCESS INBOX
    rv, data = M.search(None, "(UNSEEN)")
    message_num = 1
    new_emails = []
    for num in data[0].split():
        rv, data = M.fetch(num, '(RFC822)')
        msg = email.message_from_bytes(data[0][1])
        hdr = email.header.make_header(email.header.decode_header(msg['Subject']))
        sender = email.header.make_header(email.header.decode_header(msg['From']))
        M.store(num, "-FLAGS", '\\SEEN')
        subject = str(hdr)
        sender = str(sender)
        mail = {"message_num": message_num, "sender": sender, "subject": subject}
        new_emails.append(mail)
        message_num += 1
    
    M.close()
    M.logout()

    return new_emails

class Email(MycroftSkill):
    def __init__(self):
        MycroftSkill.__init__(self)

    @intent_file_handler('check.email.intent')
    def handle_email(self, message):
       #Get settings
       account = self.settings.get('username')
       server = self.settings.get("server")
       if account == None or account == "" or server == None or server == "":
           self.speak_dialog("setup")
           return
       folder = self.settings.get('folder')
       password = self.settings.get('password')
       port = self.settings.get("port")
       #check email
       new_emails = list_new_email(account=account, folder=folder, password=password, port=port, address=server)
       
       #report back
       for x in range(0, len(new_emails)):
           new_email = new_emails[x]
           self.speak_dialog("list.subjects", data=new_email)

def create_skill():
    return Email()

