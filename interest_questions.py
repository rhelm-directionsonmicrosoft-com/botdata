# interest_questions.py
# Extract customer questions from Interests email files
# Writes to CSV file

import os, os.path
import csv
import extract_msg as msg
import email
import email.parser
import email.policy
import re


inpath = r"/Users/robhelm/dev/data/botdata/interests"
outpath  = r"/Users/robhelm/dev/data/botdata/interests/interests_out.csv"
fields = ['sender','to','subject','question']
outmessage = {
    'sender':'',
    'to':'',
    'subject':'',
    'question':'',
    }


# Get e-mail messages from a folder and merge into a CSV file.

def get_messages(inpath=inpath, outpath=outpath):
    with open(outpath, 'w', newline='', encoding='utf-8') as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=fields, dialect='excel')
        writer.writeheader()
        for message in read_messages(inpath):
            print(f'{message["subject"]}')
            writer.writerow(message)
    csv_file.close()

# Read messages in folder and return them to get_messages

def read_messages(inpath):
    for filename in os.scandir(inpath):
        message = read_message(filename)
        if target_message(message):
            print(f'read target sender {message["sender"]}\nsubject {message["subject"]} ') # DBG
            yield(message)
        else:
            print('read_messages no target message') # DBG

# read message file, identify format, and parse to message

def read_message(filename):
    path = filename.path
    ext = message_file_type(filename)
    message = dict(outmessage)
    try:    
        if ext == '.msg': # Outlook binary format
            message_source = msg.openMsg(path)
        elif ext == '.eml': # Quasi-standard SMTP/IMAP format
            message_source = EmlMessageSource(path)
        else:
            # DBG print(f'Not a known message format')
            message_source = None
        if message_source:
            message['sender'] = message_source.sender
            print(f'read_source sender {message_source.sender}\nsubject: {message_source.subject}')
            message['to']= message_source.to
            message['subject'] = message_source.subject
            message['question'] = extract_question(message_source.body)
            print(f'read question {message["question"]}')
    finally:
        if message_source:
            message_source.close()
            # print(f'read_message: {message["subject"]}') # DBG
        else:
            pass
            # print('read_message: none') #DBG
    return(message)

def target_message(message):
    return(True) #DBG
    subject = message.get('subject', None)
    if subject == None:
        return(False)
    else:
        m = re.match(subject_pat,subject)
        return(m)

subject_pat = re.compile(r'Customer\s+Query\s+from(.*)$', re.IGNORECASE)

question_pat = re.compile(r'in\s+Query(.*?)[-_]{2}', re.IGNORECASE)

clean_pat = re.compile(r'\s\s|\S\n\S')

def clean():
    return(' ')


def extract_question(message):
    # body = re.sub(clean_pat, clean, message['body']) #DBG
    m = re.search(question_pat)
    if not m:
        return(None)
    else:
        question = m.group(1) # m.group('question') #DBG
        return(question)
        
    
def message_file_type(filename):
    path = filename.path
    if os.path.isfile(path):
        (base,extension) = os.path.splitext(path)
        if extension == '.msg' or extension == '.eml':
            # print(f'message_file_type: {extension}') DBG
            return(extension)
        else:
            # print('message_file_type: not a message file type') DBG
            return('')
    else:
        # print('message_file_type: not a file') DBG
        return('')
    
class EmlMessageSource():
    
    def __init__(self, path):
        self.eml_parser = email.parser.BytesParser(policy=email.policy.default)
        with open(path, mode='rb') as eml_file:
            message = eml_parser.parse(eml_file)
        eml_file.close()
        self.message = message
        self.subject = message['subject']
        print('EmlMessage source init subject:{message["subject"]}\nsender: {message.sender}')
        self.to = message['to']
        self.sender =  message['sender']
        self.body = self.get_eml_body(message)

    def close():
        pass

            
    # Get the text body from an eml message
    def get_eml_body(self, message):
        part = message.get_body(preferencelist=('plain', 'html'))
        content_type = part['content-type']
        if (not content_type.maintype == 'text'):
            # print('Not text format, blank body returned')  # DBG
            body = ' '
        else:
            body = part.get_content()
            # DBG print(f'body format{content_type.maintype}\\{content_type.subtype}')#DBG
        return(body)




'''
if __name__ == '__main__':
    result = get_messages()
'''
