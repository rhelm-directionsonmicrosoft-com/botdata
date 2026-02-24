# load_questions.py
# Extract customer questions from email and Excel files
# Writes to json file via Pandas DataFrame

import os, os.path
import pathlib
import pandas as pd
import extract_msg as msg
import email
import email.parser
import email.policy
import re
from random import randint


inpath = r"C:\Users\Rob\script\data\botdata_files\interests"
outpath  = r"C:\Users\Rob\script\data\botdata_files\interests\interests_out.csv"
outscript  = r"C:\Users\Rob\script\data\botdata_files\interests\interests_script.txt"
fields = ['sender','to','subject','question']


    
class EmlSource():
    outmessage = {
    'sender':'',
    'to':'',
    'subject':'',
    'question':'',
    }
    
    def __init__(self):
        self.eml_parser = email.parser.BytesParser(policy=email.policy.default)
        self.message = None
        self.subject = ''
        self.to = ''
        self.sender = ''
        self.body = ''
        # print('EmlSource {self} initialized') #DBG

    def open_eml(self, path):
        with open(path, mode='rb') as eml_file:
            message = self.eml_parser.parse(eml_file)
        eml_file.close()
        self.message = message
        self.subject = message['subject']
        self.to = message['to']
        self.sender =  message['sender']
        self.body = self.get_body(message)
        # print('eml_source.open subject:self.subject}\nsender: {message.sender}') # DBG
        return(self)

    def close(self):
        pass
            
    # Get the text body from an eml message
    def get_body(self, message):
        part = message.get_body(preferencelist=('plain', 'html'))
        content_type = part['content-type']
        if (not content_type.maintype == 'text'):
            # print('Not text format, blank body returned')  # DBG
            body = ' '
        else:
            body = part.get_content()
            # print(f'body format{content_type.maintype}\\{content_type.subtype}') #DBG
        return(body)

eml_source = EmlSource() # Reader for .eml messages

# Get e-mail messages from a folder and merge into a CSV file.

# write out scripts for the bot

def write_bracketed(
    inpath=inpath,
    outpath=outscript,
    delay=120*randint(1, 5),
    MicrosoftLearn=True):
    with open(outpath, mode='w',encoding='utf-8') as f:
        print('### Interests e-mailed member questions from 2024 and earlier',file=f)
        for question in read_questions(inath):
            if MicrosoftLearn == True:
                print('USE Microsoft Learn',file=f)
            print('QUESTION{{' + question.strip() + '}}',file=f)
            print('DELAY{{' + str(delay) + '}}',file=f)
        print('### Done',file=f)
    f.close()
                      

    
def load_data(inpath=inpath):
    with open(outpath, 'w', newline='', encoding='utf-8') as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=fields, dialect='excel')
        writer.writeheader()
        for message in read_messages(inpath):
            # print(f'{message["subject"]}')
            writer.writerow(message)
    csv_file.close()

# Read questions in a source folder or directory and return them as read

def read_questions(inpath):
    for filename in os.scandir(inpath):
        doc = read_doc(filename)
        if target_doc(doc):
            # print(f'read target sender {message["sender"]}\nsubject {message["subject"]} ') # DBG
            yield(message)
        else:
            pass
            # print('read_messages no target message') # DBG

# read message file, identify format, and parse to message

def read_message(filename):
    path = filename.path
    ext = message_file_type(filename)
    message = dict(outmessage)
    try:    
        if ext == '.msg': # Outlook binary format
            message_source = msg.openMsg(path)
        elif ext == '.eml': # Quasi-standard SMTP/IMAP format
            message_source = eml_source.open_eml(path)
        else:
            # DBG print(f'Not a known message format')
            message_source = None
        if message_source:
            message['sender'] = message_source.sender
            # print(f'read_source sender {message_source.sender}\nsubject: {message_source.subject}') #DBG
            message['to']= message_source.to
            message['subject'] = message_source.subject
            message['question'] = '' # default value if not found in body
            body = message_source.body
            if body:
                # re.sub(clean_pat, cleaner, message_source.body) #DBG
                if body:
                    #question = extract_question(body)
                    question = body
                    message['question'] = question
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

question_pat = re.compile(r'.*', re.IGNORECASE | re.MULTILINE | re.DOTALL)
# = re.compile(r'(?:Member Question|in Query)[:](.*?)[-_][-_]|\n\n', re.IGNORECASE | re.MULTILINE)

clean_pat = re.compile(r'\s\s|\S\n\S')

def cleaner(m):
    return(' ')


def extract_question(body):
    question = '' # default value if not found
    m = re.search(question_pat, body)
    if not m:
        print('search failed')
    elif len(m.groups()) < 1:
        for k, g in enumerate(m.groups()):
            print(f'Lacks group {k}: {g}')
    else:
        question = m.group(0) # m.group('question') #DBG
        for k, g in enumerate(m.groups()):
            print(f'Lacks group {k}: {g}')
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



'''
if __name__ == '__main__':
    result = get_messages()
'''
