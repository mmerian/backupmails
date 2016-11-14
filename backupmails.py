#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Simple script for backing up all emails from an IMAP server
Usage : python backupmails.py --type=imap --host=<imap server> --port=<imap port> --username=<imap login> --password=<imap password>
Password is asked if not provided by the command line
"""

import argparse
import sys
import email
import mailbox
import re

from getpass import getpass
from imaplib import IMAP4

parser = argparse.ArgumentParser(description = 'Check IMAP server')

parser.add_argument('--type', help='mail server type (imap or pop, default imap)', default = 'imap')
parser.add_argument('--host', help = 'mail server host name', required = True)
parser.add_argument('--port', help = 'mail server port', default = 143)
parser.add_argument('--username', help = 'mail server username', required = True)
parser.add_argument('--password', help = 'mail server password', required = False, default = None)
parser.add_argument('--mboxprefix', help = 'mailbox filename prefix', required = False, default = '')

args = parser.parse_args()

client = None
mbox = None

list_response_pattern = re.compile(r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?P<name>.*)')

def backup_imap_folder(folder):
    print folder
    try:
        client.select(folder, True)
        r, data = client.uid('search', None, 'ALL')
        mbox = mailbox.mbox(args.mboxprefix+folder+'.mbox')
        mbox.lock()
        mbox.clear()
        for msgid in data[0].split():
            try:
                r, msgdata = client.uid('fetch', msgid, '(RFC822)')
                message = email.message_from_string(msgdata[0][1])
                mbox.add(message)
            except Exception as e:
                print e
                pass
        mbox.flush()
        mbox.close()
    except Exception as e:
        print e
        pass
try:
    password = args.password
    if password is None:
        password = getpass('password : ')
    client = IMAP4(args.host, args.port)
    client.login(args.username, password)
except Exception as e:
    print(e)
    sys.exit(2)

r, response = client.list()
if 'OK' == r:
    for line in response:
        flags, delimiter, mailbox_name = list_response_pattern.match(line).groups()
        try:
            backup_imap_folder(mailbox_name)
        except Exception as e:
            print e
            pass
