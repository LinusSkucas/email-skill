# Copyright 2018 Linus S
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
# 
#    http://www.apache.org/licenses/LICENSE-2.0
# 
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#
from mycroft import MycroftSkill, intent_file_handler
import sys
import imaplib
import email
import email.header

def list_new_email(account, folder, password, port, address):
    """
    Returns new emails.
    output:
    dicts in the format :{name,sender,subject}
    """
    M = imaplib.IMAP4_SSL(str(address), port=int(port))
    M.login(str(account), str(password))#Login
    M.select(str(folder))
    rv, data = M.search(None, "(UNSEEN)")#Only get unseen/unread emails
    message_num = 1
    new_emails = []
    for num in data[0].split():
        rv, data = M.fetch(num, '(RFC822)')
        msg = email.message_from_bytes(data[0][1])
        hdr = email.header.make_header(email.header.decode_header(msg['Subject']))#Get subject
        sender = email.header.make_header(email.header.decode_header(msg['From']))#Get sender
        M.store(num, "-FLAGS", '\\SEEN') #Some email providers automaticly mark a message as seen: undo that.
        subject = str(hdr)
        sender = str(sender)
        mail = {"message_num": message_num, "sender": sender, "subject": subject}
        new_emails.append(mail)
        message_num += 1
    #Clean up
    M.close()
    M.logout()

    return new_emails

class Email(MycroftSkill):
    """The email skill
    Checks your emails for you
    """
    def __init__(self):
        """Init"""
        MycroftSkill.__init__(self)

    @intent_file_handler('check.email.intent')
    def handle_email(self, message):
       """Get the new emails and speak it"""
       #Get settings on home
       account = self.settings.get('username')
       server = self.settings.get("server")
       password = self.settings.get("password")
       if account == None or account == "" or server == None or server == "":
           config = self.config_core.get("email_login", {})
           account = config.get("email")#Get settings in config file
           password = config.get("password")
           if account == None or account == "" or password == "" or password == None:
               #Not set up in file/home
               self.speak_dialog("setup")
               return
       folder = self.settings.get('folder')
       port = self.settings.get("port")
       #check email
       try:
           new_emails = list_new_email(account=account, folder=folder, password=password, port=port, address=server)
       except Exception as e:
           #Error? give an error
           self.speak_dialog("error.getting.mail")
           return
       if len(new_emails) == 0:
           #No email? Say that.
           self.speak_dialog("no.new.email")
           return

       stop_num = 10
       #report back
       for x in range(0, len(new_emails)):
           new_email = new_emails[x]
           self.speak_dialog("list.subjects", data=new_email)
           #Say 10 emails, if more ask if user wants to hear them
           if x == stop_num:
               more = self.ask_yesno(prompt="more.emails")
               if more == "no":
                   self.speak_dialog("no.more.emails")
                   break
               elif more == "yes":
                   stop_num += 10
                   continue


def create_skill():
    return Email()

