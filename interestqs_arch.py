# interestqs.py
# Extract customer questions from Interests email files
# Writes to CSV file

import os, os.path
import pandas as pd

import datetime
import json
import eml_parser

# Sample file

path = r'C:\Users\Rob\script\botdata\interests\Azure Stack HCI Branch Office Licensing Discusion - Copy.eml'

# Convert time in a json file to iso

def json_serial(obj):
  if isinstance(obj, datetime.datetime):
      serial = obj.isoformat()
      return serial

# Dump a .eml file to json on stdout

def dump_json(path=path):
    print(f'opening {path}')
    with open(path, 'rb') as fhdl:
        raw_email = fhdl.read()
    fhdl.close()
    print(f'raw email read')
    ep = eml_parser.EmlParser()
    print('EMLParser loaded')
    parsed_eml = ep.decode_email_bytes(raw_email)
    print(json.dumps(parsed_eml, default=json_serial))

if __name__ == '__main__':
    dump_json()
