# interestqs.py
# Extract customer questions from Interests email files
# Writes to CSV file

import os, os.path
import csv
import extract_msg as msg
import email
import email.policy
import email.parser
from pprint import pprint as pp
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
        self.eml_parser = email.parse.EmlParser()
        with open(path, mode='rb') as eml_file:
            raw_email = eml_file.read()
        # Handle multipart message  TODO
        # decode and return TODO
    def close(self):
        pass
    # handle dict read TODO      
    

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
        elif ext == '.eml':
                message_source = EmlMessageSource(path)

        else:
            print(f'Not a message file')
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
