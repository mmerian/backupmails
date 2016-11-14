#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Simple script for backing up all emails from an IMAP server.
The script create one mailbox file for each IMAP folder

Usage : python backupmails.py --host=<imap server> --port=<imap port> --username=<imap login> --password=<imap password>

Password is asked if not provided by the command line
"""

import os
import argparse
import sys
import email
import mailbox
import re

from getpass import getpass
from imaplib import IMAP4,IMAP4_SSL

parser = argparse.ArgumentParser(description = 'Creates mailbox files from messages on an IMAP server')

parser.add_argument('--host', help = 'mail server host name', required = True)
parser.add_argument('--port', help = 'mail server port', type=int, default = 143)
parser.add_argument('--username', help = 'mail server username', required = True)
parser.add_argument('--password', help = 'mail server password', required = False, default = None)
parser.add_argument('--mboxprefix', help = 'mailbox filename prefix', required = False, default = '')
parser.add_argument('--ssl', help = 'use SSL connection', action='store_true')
parser.add_argument('--dest-dir', help = 'Mailbox files destination directory', required = False, default = None)

args = parser.parse_args()

client = None
mbox = None

list_response_pattern = re.compile(r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?P<name>.*)')

def backup_imap_folder(folder):
    print folder
    try:
        client.select(folder, True)
        r, data = client.uid('search', None, 'ALL')
        mboxfilename = args.mboxprefix+folder+'.mbox'
        if args.dest_dir is not None:
            mboxfilename = os.path.join(args.dest_dir, mboxfilename)
        mbox = mailbox.mbox(mboxfilename)
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

# Entry point
try:
    # Check for password
    password = args.password
    if password is None:
        password = getpass('password : ')

    # If port is 993, then turn on SSL
    ssl = args.ssl
    if 993 == args.port:
        ssl = True

    # Login to IMAP server
    if ssl:
        client = IMAP4_SSL(args.host, args.port)
    else:
        client = IMAP4(args.host, args.port)
    client.login(args.username, password)

    # Create dest dir if not exists
    if args.dest_dir is not None:
        if not os.path.isdir(args.dest_dir):
            os.mkdir(args.dest_dir)

except Exception as e:
    print(e)
    sys.exit(2)

# Iterate through all IMAP folders and backup messages
r, response = client.list()
if 'OK' == r:
    for line in response:
        flags, delimiter, mailbox_name = list_response_pattern.match(line).groups()
        try:
            backup_imap_folder(mailbox_name)
        except Exception as e:
            print e
            pass
