from mycroft import MycroftSkill, intent_file_handler
from check_email import list_new_email

class Email(MycroftSkill):
    def __init__(self):
        MycroftSkill.__init__(self)

    @intent_file_handler('check.email.intent')
    def handle_email(self, message):
       #Get settings
       first_run = self.settings.get('first_run')
       if first_run == None:
           self.speak_dialog("setup")
           return
       account = self.settings.get('username')
       folder = self.settings.get('folder')
       password = self.settings.get('password')
       port = self.settings.get("port")
       server = self.settings.get('server')
       #check email
       new_emails = list_new_email(account=account, folder=folder, password=password, port=port, server=server)
       #report back
       for x in range(0, len(new_emails)):
           self.speak_dialog("list.subjects", data=new_emails[x])

def create_skill():
    return Email()

