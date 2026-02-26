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

# Abstract class gets questions from files

class QuestionSource()

    # True if filename extension can be processed here
    # SUBCLASS RESPONSIBILITY
    def myextension(self, ext):
        return(False)

    # Read doc with questions
    # SUBCLASS RESPONSIBILITY
    def read(self, filepath):
        pass

        
    def __init__(self, folderpath=inpath, fields={}):
        self.folderpath = folderpath
        self.parser = None
        self.fields = fields.copy()
        self.message = self.fields.copy()
        self.message['question'] = ''


    # Get question from body text of a source file

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
        message = self.clean(message)
        m = re.search(question_pat, message['body'])
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

    # clean up message body before scanning for question

    clean_pat = re.compile(r'\s\s|\S\n\S')

    def cleaner(m):
        return(' ')

    def clean(self, message):
        body = re.sub(clean_pat, cleaner, message['body'])
        message['body'] = body
        return(message)


# get questions from .eml files

    
class EmlSource(QuestionSource):

    # True if filename extension can be processed here

    def myextension(self, ext):
        return(ext == '.eml')
  
    def __init__(self, path, fields):
        QuestionSource.__init__(self, path, fields)
        self.parser = email.parser.BytesParser(policy=email.policy.default)
        # print('EmlSource {self} initialized') #DBG

    # Read a message with question from .eml file

    def read(self, filepath):
        with open(self.filepath, mode='rb') as eml_file:
            _message = self.eml_parser.parse(eml_file)
        eml_file.close()
        message = self.message.copy()
        for field in field.keys():
            message[field] = self.get_field(_message, field)
        message['question'] = self.get_question(message)
        # print('eml_source.open subject:self.subject}\nsender: {message.sender}') # DBG
        return(message)

    # Get a field from message being processed.
    
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
    

# Get questions from folder of .msg files


class MsgSource(QuestionSource):
        # True if filename extension can be processed here

    def myextension(self, ext):
        return(ext == '.msg')
    
    def __init__(self, path, fields):
        QuestionSource.__init__(self, path)
        self.fields = fields.copy()
        self.message = fields.copy()
        self.message['question'] = ''

    def read(self, filepath):
        _message = msg.openMsg(filepath)
        for field in self.fields.keys():
            message[field] = _message[field]
        message['question'] = self.get_question(message)
        return(message)

# TODO 2026-02-25
            
class CsvSource(QuestionSource):
    
        # True if filename extension can be processed here

    def myextension(self, ext):
        return(ext == '.csv')
    
    __init(self, path, fields)__:
        QuestionSource.__init__(self, path, fields)
        self.buffer = pd.DataFrame(filepath, columns = list(self.message.keys()))

    def read(self, filepath):
        with pd.ExcelFile(filepath) as xlfile:
            self.buffer = concat_worksheets(xlfile, filepath, self.buffer
        xlfile.close()
        
                                            
    def concat_worksheets(xlfile, filepath, buffer):
        sheet_names = xlfile.sheet_names
        for n in sheet_names:
            print(f'concat_worksheet {n}')
            new_sheet = pd.DataFrame = pd.read_excel(xlfile, n)
            new_sheet.name = n
            new_sheet.insert(len(new_sheet.columns), "source_path", str(path))
            new_sheet.insert(len(new_sheet.columns), "source_sheet", n)
            result = pd.concat([result, new_sheet], ignore_index=True)
        print(f'concat_worksheets {result.info()}')

    def read(self):
        

        
    
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



    

    # Read all question files in a folder, return one by one

    def read_questions():
        for filename in os.scandir(self.path):
            doc = read_doc(filename)
            if target_doc(doc):
                yield(doc)

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

# Read questions in a source folder and return them as read



'''
if __name__ == '__main__':
    result = get_messages()
'''
