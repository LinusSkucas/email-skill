from mycroft import MycroftSkill, intent_file_handler


class Email(MycroftSkill):
    def __init__(self):
        MycroftSkill.__init__(self)

    @intent_file_handler('email.intent')
    def handle_email(self, message):
        self.speak_dialog('email')


def create_skill():
    return Email()

