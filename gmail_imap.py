import imaplib
import email
import os

class CImap():
    """this class handles various imap operations"""

    def __init__(self, username,password, label):
        self.attachments = []
        self.res_connection = imaplib.IMAP4_SSL('imap.gmail.com')
        try:
            self.res_connection.login(username, password)
        except Exception as e:
            raise e

        self.res_connection.select(label)

    def close_connection(self):
        self.res_connection.close()

    def search_emails( self, criteria = 'All'):
        self.type, self.message_nums = self.res_connection.search(None, criteria)
        self.message_nums = self.message_nums[0].split()

    def get_message_number(self):
        return self.message_nums

    def get_email_details(self, message_num):
        self.email = self.res_connection.fetch(message_num, '(RFC822)')

    def delete_email(self, message_num):
        self.res_connection.store(num, '+FLAGS', '\\Deleted')
        ObjImap.res_connection.expunge()

    def download_attachments(self, download_path=None):
        raw_email = self.email[1][0][1]
        raw_email_string = raw_email.decode('utf-8')
        email_message = email.message_from_string(raw_email_string)

        for part in email_message.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            file_name = part.get_filename()

            if True == bool(file_name):
                if None == download_path:
                    download_path = os.getcwd()

                download_path = download_path + '\\' + file_name
                self.attachments.append(file_name)

                fp = open(download_path, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
        return True

username = 'test'
password = 'test'
ObjImap = CImap(username, password, 'Inbox')
ObjImap.search_emails('(FROM "test@gmail.com" SUBJECT "test" )')
message_num = ObjImap.get_message_number()
for num in message_num:
    ObjImap.get_email_details(num)
    #ObjImap.delete_email(message_num)
    ObjImap.download_attachments()
    print('Downloaded attachments')
    print(ObjImap.attachments)

ObjImap.close_connection()
