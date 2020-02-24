# Copyright 2019 Linus S
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
from mycroft.filesystem import FileSystemAccess
from mycroft.skills.settings import save_settings
from datetime import datetime
import sys
import imaplib
import email
import email.header
import email.utils
import inflect
import yaml

EMAIL_POLL_INTERVAL = 20  # in seconds


def normalize_email(email):
    if not email:
        return None

    result = ""

    for token in email.split():
        if token == "dot":
            result += "."
        elif token == "at":
            result += "@"
        else:
            result += token

    return result


class Email(MycroftSkill):
    """The email skill checks your emails for you
    """

    def __init__(self):
        """Init"""
        MycroftSkill.__init__(self)

        self.i_engine = inflect.engine()

    def update_credentials(self):
        self.server = False  # False for an ics file, true for a caldav server -regardless of where the creds are stored
        self.no_creds = False

        self.server = True

        self.account = self.settings.get('username')
        self.password = self.settings.get('password')
        self.server = self.settings.get("server")
        self.folder = self.settings.get('folder')
        self.port = self.settings.get("port")

        if self.account is None or self.account == "" or self.server is None or self.server == "":
            # Using pw in config
            fs = FileSystemAccess(str(self.skill_id))
            if fs.exists("email_conf.yml"):
                #  Use yml file for config
                config = fs.open("email_conf.yml", "r").read()
                config = yaml.safe_load(config)
                self.account = config.get("username")
                self.password = config.get("password")
                self.server = config.get("server_address")
                self.folder = config.get("folder")
                self.port = config.get("port")

            else:
                self.no_creds = True
            if self.account is None or self.account == "":
                self.no_creds = True

        if self.no_creds is True:
            # Not set up in file/home
            self.speak_dialog("setup")
            return False
        return True

    def initialize(self):
        # Start the notification service if it was active when Mycroft quit
        self.remove_event('poll.emails')
        if self.settings.get('look_for_mail'):
            self.schedule_repeating_event(self.poll_emails, datetime.now(), EMAIL_POLL_INTERVAL, name='poll.emails')

        self.update_credentials()

    def _nice_number(self, num):
        return self.i_engine.number_to_words(self.i_engine.ordinal(num))

    def report_email(self, new_emails):
        """Report and say the email"""
        stop_num = 10
        x = 0
        # report back
        if not new_emails:
            return

        for idx, new_email in enumerate(list(reversed(new_emails)), start=1):  # newest to oldest
            if self.lang == "en-us":  # TODO: localize this!
                new_email["message_num"] = self._nice_number(idx)
            else:
                pass
            self.speak_dialog("list.subjects", data=new_email)
            x += 1
            # Say 10 emails, if more ask if user wants to hear them
            if x == stop_num:
                more = self.ask_yesno(prompt="more.emails")
                if more == "no":
                    self.speak_dialog("no.more.emails")
                    break
                elif more == "yes":
                    stop_num += 10
                    continue

    def list_new_email(self, account, folder, password, port, address, whitelist=None, mark_as_seen=False):
        if self.update_credentials() is False:  # No credentials
            return
        """
        Returns new emails.
        output:
        dicts in the format :{name,sender,subject}
        whitelist: the a list of emails that it will return
        """
        M = imaplib.IMAP4_SSL(str(address), port=int(port))
        M.login(str(account), str(password))  # Login
        M.select(str(folder))
        rv, data = M.search(None, "(UNSEEN)")  # Only get unseen/unread emails
        message_num = 1
        new_emails = []
        for num in data[0].split():
            rv, data = M.fetch(num, '(RFC822)')
            msg = email.message_from_bytes(data[0][1])
            hdr = email.header.make_header(email.header.decode_header(msg['Subject']))  # Get subject
            sender = email.header.make_header(email.header.decode_header(msg['From']))  # Get sender
            from_email = email.utils.parseaddr(msg['From'])[1]
            subject = str(hdr)
            sender = str(sender).lower()

            is_in_whitelist = not whitelist or from_email in whitelist or any(
                s.lower() in sender.lower().split(" ") for s in whitelist)
            if is_in_whitelist and mark_as_seen:
                M.store(num, "+FLAGS", '\\SEEN')
            else:
                M.store(num, "-FLAGS", '\\SEEN')  # Some email providers automatically mark a message as seen: undo that
            # For english ordinals TODO: I should localize this
            mail = {"message_num": message_num, "sender": sender, "subject": subject}
            if not is_in_whitelist:
                # The user does not want emails from that sender, skip it
                continue

            new_emails.append(mail)
            message_num += 1
        # Clean up
        M.close()
        M.logout()

        return list(new_emails)

    def poll_emails(self, data):
        setting = self.settings.get('look_for_email')
        # check email
        try:
            new_emails = self.list_new_email(account=self.account, folder=self.folder, password=self.password,
                                             port=self.port, address=self.server, whitelist=setting['whitelist'],
                                             mark_as_seen=True)
        except Exception as e:
            # Silently ignore errors
            return

        if not new_emails:
            # No new mail
            return

        stop_num = 10
        num_emails = len(new_emails)
        response = self.ask_yesno(prompt="notify.read.email", data={"size": num_emails})
        if response != "yes":
            return

        self.report_email(new_emails)

    @intent_file_handler('enquire.email.intent')
    def enquire_new_email(self, message):
        if self.update_credentials() is False:  # No credentials
            return
        # get the sender
        sender = message.data.get('email').lower()
        setting = self.settings.get('look_for_email')
        # check email
        try:
            new_emails = self.list_new_email(account=self.account, folder=self.folder, password=self.password,
                                             port=self.port, address=self.server, whitelist=[sender], mark_as_seen=True)
        except Exception as e:
            # Silently ignore errors
            return

        if not new_emails:
            # No new mail
            self.speak_dialog("no.emails.from", data={'sender': sender})
            return

        self.report_email(new_emails)

    @intent_file_handler('notify.intent')
    def enable_email_polling(self, message):
        if self.update_credentials() is False:  # No credentials
            return
        sender = normalize_email(message.data.get('sender'))
        setting = self.settings.get('look_for_email')
        if not setting:
            # The email notification service is not active yet
            if sender:
                self.settings['look_for_email'] = {"whitelist": [sender]}
            else:
                self.settings['look_for_email'] = {"whitelist": None}

            # Start the notification service
            self.schedule_repeating_event(self.poll_emails, datetime.now(), EMAIL_POLL_INTERVAL, name='poll.emails')
            self.speak_dialog("start.poll.email")
        else:
            if not setting['whitelist'] and sender:
                # The user request that we look for a specific email, but we are already notifying all incoming mails
                # Ask if the user wants to limit the notifications to that email
                replace = self.ask_yesno(prompt="cancel.looking.for.all.look.specific", data={"email": sender})
                if replace == "yes":
                    setting['whitelist'] = [sender]
                else:
                    # We did not change anything, so the event data does not need to be updated
                    return
            elif not setting['whitelist'] and not sender:
                # We were requested to look for all emails, but we are already doing that
                self.speak_dialog("already.looking.for.all")
                return
            elif sender in setting['whitelist']:
                # We were requested to look for emails by a specific person, but we are already doing that
                self.speak_dialog("already.looking.for.specific", data={"email": sender})
                return
            elif setting['whitelist'] and not sender:
                replace = self.ask_yesno(prompt="cancel.looking.for.specific.look.all",
                                         data={"email": setting['whitelist']})
                if replace == "yes":
                    setting['whitelist'] = None
                else:
                    # We did not change anything, so the event data does not need to be updated
                    return
            else:
                setting['whitelist'].append(sender)

            self.speak_dialog("update.notify.data")
        save_settings(self.root_dir, self.settings)

    @intent_file_handler('stop.intent')
    def disable_email_polling(self, message):
        if not self.settings.get('look_for_email'):
            self.speak_dialog("poll.emails.not.started")
            return

        sender = normalize_email(message.data.get('sender'))

        if not sender:
            self.settings['look_for_email'] = None
            self.remove_event('poll.emails')
            self.speak_dialog("stop.poll.emails")
        else:
            # We are looking for all new emails, turn it off completely
            if not self.settings['look_for_email']['whitelist']:
                self.settings['look_for_email'] = None
                self.remove_event('poll.emails')
                self.speak_dialog("stop.poll.emails")
                save_settings(self.root_dir, self.settings)
                return

            # Do we even look for that sender?
            if not sender in self.settings['look_for_email']['whitelist']:
                self.speak_dialog("not.looking.for", data={"email": sender})
                return

            if len(self.settings['look_for_email']['whitelist']) > 1:
                # Remove the sender from the whitelist
                self.settings['look_for_email']['whitelist'].remove(sender)
                self.speak_dialog("stop.looking.for", data={"email": sender})
            else:
                # We don't have any email addresses to look for, so turn it off completely
                self.settings['look_for_email'] = None
                self.remove_event('poll.emails')
                self.speak_dialog("stop.poll.emails.last.email.removed")

        save_settings(self.root_dir, self.settings)

    @intent_file_handler('check.email.intent')
    def handle_email(self, message):
        """Get the new emails and speak it"""
        self.log.info("Checking email...")
        # check email
        try:
            new_emails = self.list_new_email(account=self.account, folder=self.folder, password=self.password,
                                             port=self.port, address=self.server)
        except Exception as e:
            # Error? give an error
            self.speak_dialog("error.getting.mail")
            return

        if not new_emails:
            # No email? Say that.
            self.speak_dialog("no.new.email")
            return

        self.report_email(new_emails)


def create_skill():
    return Email()
