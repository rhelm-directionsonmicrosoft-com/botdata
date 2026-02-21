# interestqs.py
# Extract customer questions from Interests email files
# Writes to CSV file

import os, os.path
import pandas as pd
import datetime
import json
import eml_parser
import extract_msg as emsg
from pprint import pprint as pp
import re


inpath = r"/Users/robhelm/dev/data/botdata/interests"
outpath  = r"/Users/robhelm/dev/data/botdata/interests/interests_out.txt"

def combine_message_files(inpath=inpath, outpath=outpath):
    result = concat_files(inpath,list([]))
    with open(outpath, mode='w', encoding='utf-8') as outf:
        pp(result, stream=outf)
    outf.close()
    return(result)

pattern = re.compile(r'\r\n')

def clean_message_body(body):
    def cleaner(matched):
        return('\n')
    body = re.sub(pattern, cleaner, body)
    return(body)
    

def concat_files(inpath, result):
    for filename in os.scandir(inpath):
        extension = message_file_type(filename)
        if extension == '.eml':
            is_eml = True
            print(f'is eml file')
            result = append_eml(filename.path, extension, result)
        elif extension == '.msg':
            print(f'is msg file')
            result = append_msg(filename.path, extension, result)
        else:
            print('is not message file')    
        print(f'result contains {len(result)} messages')
    return(result)

def append_eml(path, extension, result):
    # print(f'append_eml {path}{extension} {result}')
    with open(path, 'rb') as eml_file:
        raw_email = eml_file.read()
    eml_file.close()
    #print(f'raw EML file read')
    ep = eml_parser.EmlParser()
    #print('EMLParser loaded')
    parsed_eml = ep.decode_email_bytes(raw_email)
    #print('EML Parsed message:')
    result.append(parsed_eml)
    return(result)

def append_msg(path, extension, result):
    try:
        message_file = emsg.openMsg(path)
        #print(f'raw MSG file read')
        msg = dict({})
        msg['body'] = message_file.body
        msg['subject'] = (message_file.subject)
        msg['receivedTime'] = (message_file.receivedTime)
        msg['rtfBody'] = (message_file.rtfBody)
        result.append(msg)
    finally:
        message_file.close()
    return(result)
    
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


if __name__ == '__main__':
    result = combine_message_files()
