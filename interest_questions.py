# interest_questions.py
# Extract customer questions from Interests email files
# Writes to CSV file

import os, os.path
import csv
import extract_msg as msg
import email
from email.parser
from email.policy
import re

fields = {
    'Question': (extract_question,  'body')
    'Subject', (lambda x: x, 'subject')
    'Sender', (lambda x: x, 'sender')
    'To', (lambda x: x, 'to')
    }

inpath = r"/Users/robhelm/dev/data/botdata/interests"
outpath  = r"/Users/robhelm/dev/data/botdata/interests/interests_out.csv"

class EmlMessageSource():
    
    def __init__(self, path):
        self.parser = email.parser.BytesParser(policy=email.policy.default)
        with open(path, mode='rb') as eml_file:
            message = BytesParser(policy=policy.default).parse(eml_file)
        eml_file.close()
        self.message = message
        self.subject = message['subject']
        self.to = message['to']
        self.sender =  message['sender']
        self.body = self.get_body(message)
            
    # Get the text body from an email
    def get_body(self, message):
        part = message.get_body(preferencelist=('plain', 'html'))
        content_type = part['content-type']
        if (not content_type.maintype == 'text'):
            print('Not text format')
            body = ' '
        else:
            body = part.get_content()
            print(f'body format{content_type.maintype}\{content_type.subtype}')#DBG
        return(body)
    

def get_messages(inpath=inpath, outpath=outpath):
    with open(outpath, 'w', newline='', encoding='utf-8') as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=columns, dialect='excel')
        writer.writeheader()
        result = list([])
        for message in read_messages:
            writer.writerow(message)
            result.append(message) #debug
    csv_file.close()
    return(result)

def read_messages(inpath, result):
    for filename in os.scandir(inpath):
        message = read_message(filename)
        yield(message)

def read_message(filename):
    path = filename.path
    ext = filename.extension
    try:    
        if ext == '.msg': # Outlook binary format
            message_source = msg.openMsg(path)
        elif ext == '.eml': # Quasi-standard SMTP/IMAP format
                message_source = EmlMessageSource(path)
        else:
            print(f'Not a known message format')
            message_source = None
        if message_source:
            message = dict({})
            for (field, (reader, source_field)) in sorted(fields.items()):
                message[k] = reader(message_source[source_field])
    finally:
        message_source.close()
    return(message)

def target_message(message):
    m = re.match(subject.pat,message.subject)
    print('Target message: {m.group(1)}')
    return(m)
 
def append_message(message, result):
    if self.target_message(message):
        message['question'] = extract_question(message['body'])
        result.append(message)
        #DBG
        if len(question) > 70:
            print(f'question: {question[:70]}')
        else:
            print(f'question: {question}')
        print(f'subject: {message.subject}')
        print(f'sender: {message.sender}')
        print(f'to: {message.to}')
    return(result)
    
def message_file_type(filename):
    path = filename.path
    if os.path.isfile(path):
        (base,extension) = os.path.splitext(path)
        if extension == '.msg' or extension == '.eml':
            # print(f'message_file_type: {extension}')
            return(extension)
        else:
            # print('message_file_type: not a message file type')
            return('')
    else:
        # print('message_file_type: not a file')
        return('')


if __name__ == '__main__':
    result = combine_message_files()
