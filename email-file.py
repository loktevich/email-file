import poplib
import email
import os
import sys
import configparser
import logging
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CFG_FILE = 'config.ini'

# Loads cofiguration from file
config = configparser.ConfigParser()
config.read(os.path.join(BASE_DIR, CFG_FILE))

SAVE_FOLDER = config.get('application', 'save_folder')
MAILSERVER = config.get('application', 'mailserver')
DELETE_PROCEEDED = config.get('application', 'delete_proceeded')
LOG_FILENAME = config.get('logging', 'filename')
LOG_FORMAT = config.get('logging', 'format')
DATEFMT = config.get('logging', 'DATEFMT')
LEVEL = int(config.get('logging', 'level'))

# Logging settings
logging.basicConfig(filename=LOG_FILENAME, format=LOG_FORMAT, datefmt=DATEFMT, level=LEVEL)


def connect_pop3_server(mailserver, user, password):
    """
    Connect to the pop3 `mailserver`
    and login with the provided `user` email and `password`.
    """
    try:
        logging.debug(f'Connecting to {mailserver}')
        pop3_connection = poplib.POP3(mailserver)
        logging.info(f'Server welcome message: {pop3_connection.getwelcome().decode("utf8")}')
        pop3_connection.user(user)
        pop3_connection.pass_(password)
        logging.debug(f'Successfully connected to {mailserver}. Login as {user}')
        return pop3_connection
    except Exception as e:
        if sys.exc_info()[0] is None:
            logging.debug("Unexpected error: value none")
        logging.exception('Got exception while connecting to pop3 server')
        logging.error(e.__doc__)
        logging.error(e.message)
    

def close_pop3_connection(pop3_connection):
    """
    Close the `pop3_connection` and release the connection object
    """
    pop3_connection.quit()
    logging.debug(f'Connection to pop3 server is closed')


def process_email(pop3_connection, msg_num, save_folder, messagesCount=0):
    """
    Downloads attachements from message No.`msg_num` and saves to `save_folder`
    """
    logging.info(f'=== Processing message No.{msg_num}/{messagesCount} msg')
    (message_response, message_lines, octets) = pop3_connection.retr(msg_num)
    message_response = message_response.decode('utf8')
    if message_response.startswith('+OK'):
        logging.debug(f'Message response: {message_response}')
        message_str = [line for line in message_lines]
        message_obj = email.message_from_bytes(b'\n'.join(message_str))
        logging.debug('Start extracting attachments')
        for message_part in message_obj.walk():
            if message_part.get_content_maintype() == 'multipart':
                continue
            if message_part.get('Content-Disposition') is None:
                continue
            filenamedecode = email.header.decode_header(message_part.get_filename())
            code = filenamedecode[0][1]
            if code != None :
                filename = filenamedecode[0][0].decode(code)
            else :
                filename = filenamedecode[0][0]
            filename = re.sub('[^\w\s-]', '.', filename).strip().lower()
            filename = filename.replace('\n','')
            filename = filename.replace('\r','')
            filename = filename.replace('\t','')
            logging.debug(f'Attachment filename: {filename}')
            file_data = message_part.get_payload(decode=True)
            save_attached_file(file_data, filename, save_folder)
            if DELETE_PROCEEDED:
                pop3_connection.dele(msg_num)
                logging.info('Message was deleted from server')
    else:
        logging.error('Error retrieving message')


def save_attached_file(file_data, filename, save_folder):
    """
    Save `file_data` with name `filename` to `save_folder`
    """
    download_dir = os.path.join(BASE_DIR, save_folder)
    filepath = os.path.join(download_dir, filename)
    os.makedirs(download_dir, exist_ok=True)
    with open(filepath, 'wb') as f:
        f.write(file_data)
    logging.info(f'Attached file was saved to {filepath}')


def download_attachments(mailserver, user, password, save_folder='tmp'):
    """
    Downloads attachements from `user` email messages on pop3 server `mailserver`
    to `save_folder` (default is `/tmp/`)
    """
    logging.info('--------- PREPARING FOR ATTACHMENTS DOWNLOADS ---------')
    pop3_connection = connect_pop3_server(mailserver, user, password)
    (messagesCount, mailboxSize) = pop3_connection.stat()
    logging.debug(f'Mailbox has {messagesCount} messages, total size {mailboxSize} bytes')
    messages_list = pop3_connection.list()
    if not messages_list[0].decode('utf8').startswith('+OK'):
        logging.error('Error downloading mail list')
        exit(1)
    for message in messages_list[1]:
        msg_num, _ = message.split()
        msg_num = int(msg_num.decode('utf8'))
        process_email(pop3_connection, msg_num, save_folder, messagesCount)
    if messagesCount > 0:
        logging.info('--------- ATTACHMENTS SUCCESSFULLY DOWNLOADED ---------')
    else:
        logging.info('-------------- THERE ARE NO NEW MESSAGES --------------')
    close_pop3_connection(pop3_connection)


if __name__=='__main__':
    if len(sys.argv) == 3:
        download_attachments(MAILSERVER, sys.argv[1], sys.argv[2], SAVE_FOLDER)
    else:
        print(" Please provide user_name and password: python email-file.py <your_user> <your_password>")
        sys.exit(0)