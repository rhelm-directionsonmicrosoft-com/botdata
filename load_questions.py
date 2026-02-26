# load_questions.py
# Extract customer questions from folders of email and Excel files
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

class QuestionSource()
    
    def __init__(self, path, fields = {}):
        self.path = path
        self.parser = None
        self.fields = fields.copy()
        self.message = self.fields.copy()
        self.message['question'] = ''

    def read(self):
        pass

    # Regular expression to get question in body text

    question_pat = re.compile(
    r'''Query)[:] # Query: at start of question
    (.*?)  # Characters in the question
    # Various markers for end of question
    [<]Draft\s*Answer #<Draft Answer
    |[-_][-_] # Repeated hyphen or underline divider
    \n\s*\n # a blank line''',
    e.VERBOSE
    | re.IGNORECASE
    | re.MULTILINE)
    
    def get_question(self,message):
        m = re.search(question_pat, body)
    if not m:
        print('search failed') #DBG
    elif len(m.groups()) < 1:
        for k, g in enumerate(m.groups()):
            print(f'Lacks group {k}: {g}') #DBG
    else:
        question = m.group(1) # m.group('question') #DBG
        if len(question) > 50:
            q = question[0:49] + '...'
        else:
            q = question
        print(f'Found question {q}')
    return(question)


    clean_pat = re.compile(r'\s\s|\S\n\S')

    def cleaner(m):
        return(' ')

    def clean(self, message):
        body = re.sub(clean_pat, cleaner, message['body'])
        message['body'] = body
        return(message)


    
class EmlSource(QuestionSource):    
    def __init__(self, path, fields):
        QuestionSource.__init__(self, path)
        self.parser = email.parser.BytesParser(policy=email.policy.default)
        # print('EmlSource {self} initialized') #DBG

    def get_field(self, _message, field):
        if field != 'body':
            return(_message[field])
        else: # Get text body of possible multipart message
            part = message.get_body(preferencelist=('plain', 'html'))
            content_type = part['content-type']
            if (not content_type.maintype == 'text'):
                # print('Not text format, blank body returned')  # DBG
                body = ' '
            else:
                body = part.get_content()
                # print(f'body format{content_type.maintype}\\{content_type.subtype}') #DBG
            return(body)

    def read(self):
        with open(self.path, mode='rb') as eml_file:
            _message = self.eml_parser.parse(eml_file)
        eml_file.close()
        message = self.message.copy()
        for field in field.keys():
            message[field] = self.get_field(_message, field)
        message['question'] = self.get_question(message)
        # print('eml_source.open subject:self.subject}\nsender: {message.sender}') # DBG
        return(message)

class MsgSource(QuestionSource):
    def __init__(self, path, fields):
        QuestionSource.__init__(self, path)
        self.fields = fields.copy()
        self.message = fields.copy()
        self.message['question'] = ''

    def read(self):
        _message = msg.openMsg(self.path)
        for field in self.fields.keys():
            message[field] = _message[field]
        message['question'] = self.get_question(message)
        return(message)

# TODO 2026-02-25
            
class CsvSource(QuestionSource):
    __init(self, path, fields)__:
        QuestionSource.__init__(self, path, fields)

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




                 Get e-mail messages from a folder and merge into a CSV file.
    
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

'''
if __name__ == '__main__':
    result = get_messages()
'''
