# pyeml.py
# Extract customer questions from Interests .eml files
# Uses python base email parser

import email.parser
import email
import os, os.path
# import pandas as pd
# import datetime
# import json
# import eml_parser
# import extract_msg as emsg
from pprint import pprint as pp
import re




class InterestsEmlParser()
    inpath_default = r"/Users/robhelm/dev/data/botdata/interests/"
    outpath_default  = r"/Users/robhelm/dev/data/botdata/interests/interests_out.txt"

    def __init__(self):
        self.parser = email.parser.BytesParser(policy=email.policy.default)

    def combine_message_files(self, inpath=inpath_default, outpath=outpath_default):
        result = concat_files(inpath, list([]))
        with open(outpath, mode='w', encoding='utf-8') as outf:
            print(result, stream=outf)
        outf.close()
        return(result)
    
def concat_files(inpath, result):
    dbg_limit = 1
    for filename in os.scandir(inpath):
        extension = message_file_type(filename)
        if extension == '.eml':
            dbg_limit = dbg_limit - 1
            if dbg_limit < 0:
                break
            print(f'is {extension} file') # DBG
            result = self.append_eml(filename.path, result)
            '''
        elif extension == '.msg':
            print(f'is msg file')
            result = append_msg(filename.path, extension, result)
            '''
        else:
            print('is not eml file')   #DBG 
        print(f'result contains {len(result)} messages') # DBG
    return(result)

    # append .eml message body_txt to question file

    def append_eml(self, message_path, result):
        print(f'append_eml {message_path}') #DBG
        with open(message_path, 'rb') as message_file:
            msg = self.parser.parse(message_file)
        message_file.close()
        print(f'msg read')
        body_text = self.get_body(msg)
        print(f'append body_text: {len(body_text)} chars') #DBG
        result.append(body_text)
        return(result)

    # Get the text body from an email
    def get_body_text(self, msg):
        part = msg.get_body(preferencelist=('html', 'plain'))
        content_type = part['content-type']
        if content_type.maintype == 'text':
            if content_type.subtype == 'plain':
                body_text = self.plain_text(part.get_content())
                print('plain text')#DBG
            elif content_type.subtype == 'plain':
                body_text = self.html_text(part.get_content())
                print('html text')#DBG
            else:
                body_text = ''
        return(body_text)

    def plain_text(self, content): #TODO
        return(content)

    def html_text(self, content): # TODO
        return(content)
    
'''
def append_msg(path, extension, result):
    try:
        message_file = emsg.openMsg(path)
        print(f'raw MSG file read')
        msg = dict({})
        msg['body'] = message_file.body #clean_message_body()
        msg['subject'] = (message_file.subject)
        msg['receivedTime'] = (message_file.receivedTime)
        msg['rtfBody'] = (message_file.rtfBody)
        result.append(msg)
    finally:
        message_file.close()
    return(result)
'''
# Get extension from filename

def message_file_type(filename):
    path = filename.path
    if os.path.isfile(path):
        (base,extension) = os.path.splitext(path)
        if extension == '.msg' or extension == '.eml':
            print(f'message_file_type: {extension}')
            return(extension)
        else:
            print('message_file_type: not a message file type')
            return('')
    else:
        print('message_file_type: not a file')
        return('')


    
class TextParser(HTMLParser):
    def __init__(self):
        self.stack = list([])
    



if __name__ == '__main__':
    result = combine_message_files()
