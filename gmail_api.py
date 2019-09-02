from __future__ import print_function
import pickle
import re
import os.path
import base64
import email
import logging
import logging.config
import json
from email.mime.text import MIMEText
from datetime import datetime
from string import ascii_uppercase
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


class GmailApi:
    """
    Initiates Gmail API for reading email messages from specified sender.
    """

    def __init__(self):
        """
        Loads logging configuration file.
        Start Gmail API initiation process.
        """
        with open('logging.conf', 'r') as lc:
            log = json.load(lc)
            logging.config.dictConfig(log)
        self.log = logging.getLogger('GmailApi')
        self.__init_api()

    def __init_api(self):
        """
        Configures service API which is defined by variable api
        :param str api. Service name
        :return service. Instance of configured service (gmail or sheets)
        """
        api = 'gmail'
        SCOPES = ['https://mail.google.com/']
        version = 'v1'

        creds = None
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists(f'creds/token_{api}.pickle'):
            with open(f'creds/token_{api}.pickle', 'rb') as token:
                creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    f'creds/credentials_{api}.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(f'creds/token_{api}.pickle', 'wb') as token:
                pickle.dump(creds, token)

        self.service = build(api, version, credentials=creds)
        self.log.debug(f'Setup completed with {api}')

    def get_receipts(self):
        """
        Iterates through 30 email messages. Then results are sent to __parse_messages() method
        which returns receipts in list.
        return list. Parsed receipts from email messages in list.
        """
        self.log.info('Getting 30 newest email messages ..')
        results = self.service.users().messages().list(userId='me', maxResults=30).execute()

        messages = results.get('messages', [])
        return self.__parse_messages(messages)

    def __parse_messages(self, messages):
        """
        From first 30 emails finds messages with selected sender email.
        Then those messages are parsed to receipt items.
        :param messages. List of email messages.
        :return nested list. List of receipts items. (Cost and purchase list)
        """
        self.log.debug('Looping through available messages')
        self.data = list()
        for message in messages:

            msg = self.service.users().messages().get(userId='me', id=message['id'], format='raw').execute()
            try:
                msg_str = str(base64.urlsafe_b64decode(msg['raw'].encode('ASCII')), 'utf-8')
            except UnicodeDecodeError:
                continue

            mime_msg = email.message_from_string(msg_str)

            last_item = ''
            receipt_items = list()
            items = list()

            try:
                sender_email = re.search('<(.*)>', mime_msg['from']).group(1)
            except AttributeError:
                continue

            if sender_email == 'noreply.code.provider@maxima.lt':
                self.log.debug('Got receipt! Now parsing .. ')
                for part in mime_msg.walk():
                    if part.get_content_type() == 'text/html':
                        payload = part.get_payload(decode=True).decode('utf-8')
                        from bs4 import BeautifulSoup
                        soup = BeautifulSoup(payload, 'lxml')
                        receipt_html = soup.findAll(lambda tag: tag.name == 'pre')
                        receipt = str(receipt_html[1]).splitlines()[1:]
                        receipt_items.append('-'.join(receipt[-2].split()[3:6]))
                        for line in receipt:
                            if line.find('Mokėti') != -1:
                                receipt_items.append('-' + line.split()[-1])
                            elif not line.endswith(' A'):
                                last_item = line
                                continue
                            elif line.startswith('Nuolaida'):
                                continue
                            else:
                                line = last_item + ' ' + line
                                items.append(' '.join(line.split()[:-2]))
                                last_item = ''
                receipt_items.append(items)
                self.data.append(receipt_items)
        self.log.debug('Parsing completed! Number of receipts {0}'.format(len(self.data)))
        return self.data

    def create_message(self, sender, to, subject, message_text):
        """
        Create a message for an email.
        Args:
        sender: Email address of the sender.
        to: Email address of the receiver.
        subject: The subject of the email message.
        message_text: The text of the email message.
        Returns:
        An object containing a base64url encoded email object.
        """
        message = MIMEText(message_text)
        message['to'] = to
        message['from'] = sender
        message['subject'] = subject

        return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}

    def send_message(self, user_id, message):
        """
        Send an email message.
        Args:
        user_id: User's email address. The special value "me"
        can be used to indicate the authenticated user.
        message: Message to be sent.
        Returns: Sent Message.
        """
        try:
            message = (self.service.users().messages().send(userId=user_id, body=message).execute())
            return message
        except errors.HttpError:
            return
