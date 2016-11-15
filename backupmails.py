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
import logging
import mailbox
import re
import signal
import socket

from getpass import getpass
from imaplib import IMAP4,IMAP4_SSL

parser = argparse.ArgumentParser(description = 'Creates mailbox files from messages on an IMAP server')

parser.add_argument('--host', help = 'mail server host name', required = True)
parser.add_argument('--port', help = 'mail server port', type = int, default = 143)
parser.add_argument('--username', help = 'mail server username', required = True)
parser.add_argument('--password', help = 'mail server password', required = False, default = None)
parser.add_argument('--mboxprefix', help = 'mailbox filename prefix', required = False, default = '')
parser.add_argument('--ssl', help = 'use SSL connection', action = 'store_true')
parser.add_argument('--dest-dir', help = 'Mailbox files destination directory', required = False, default = None)
parser.add_argument('--timeout', help = 'connection timeout', type = float, default = None)
parser.add_argument('--continue', help = 'keep existing mailbox files and append new messages', action = 'store_true', dest = 'cont')

args = parser.parse_args()

client = None
mbox = None
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
logger.addHandler(logging.StreamHandler())

list_response_pattern = re.compile(r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?P<name>.*)')

def backup_imap_folder(folder):
    global mbox
    global logger
    logger.info('Processing folder '+folder)
    try:
        client.select(folder, True)
        r, data = client.uid('search', None, 'ALL')
        mboxfilename = args.mboxprefix+folder.replace('"', '').replace('/', '.')+'.mbox'
        if args.dest_dir is not None:
            mboxfilename = os.path.join(args.dest_dir, mboxfilename)
        mbox = mailbox.mbox(mboxfilename)
        mbox.lock()
        if not args.cont:
            mbox.clear()
            mbox.flush()
        msgids = data[0].split()
        nummsgs = len(msgids)
        if 1 == nummsgs:
            logger.info('1 message in folder')
        else:
            logger.info(str(nummsgs)+' messages in folder')
        processed = 0
        for msgid in msgids:
            try:
                r, msgdata = client.uid('fetch', msgid, '(RFC822)')
                message = email.message_from_string(msgdata[0][1])
                if args.cont and message.has_key('Message-ID'):
                    for m in mbox:
                        if (m.has_key('Message-ID')) and (message['Message-ID'] == m['Message-ID']):
                            logger.info('Message '+message['Message-ID']+' already in mailbox')
                            continue
                mbox.add(message)
            except Exception as e:
                logger.exception(e)
            processed += 1
            if 0 == processed % 50:
                logger.info(str(processed)+'/'+str(nummsgs)+' messages processed')
        mbox.flush()
        mbox.unlock()
        mbox.close()
    except Exception as e:
        logger.exception(e)

def handle_signal(signal, frame):
    global mbox
    global logger
    logger.info('Received signal. Exiting')
    mbox.flush()
    mbox.unlock()
    mbox.close()
    sys.exit(0)

signal.signal(signal.SIGINT, handle_signal)

# Entry point
try:
    # Set socket timeout
    if args.timeout is not None:
        socket.setdefaulttimeout(args.timeout)
    # Check for password
    password = args.password
    if password is None:
        password = getpass('password : ')

    # If port is 993, then turn on SSL
    ssl = args.ssl
    if 993 == args.port:
        ssl = True

    # Login to IMAP server
    logger.info('Connecting to mail server')
    if ssl:
        client = IMAP4_SSL(args.host, args.port)
    else:
        client = IMAP4(args.host, args.port)
    logger.info('Logging in')
    client.login(args.username, password)
    logger.info('Logged in')

    # Create dest dir if not exists
    if args.dest_dir is not None:
        if not os.path.isdir(args.dest_dir):
            os.mkdir(args.dest_dir)

except Exception as e:
    logger.exception(e)
    sys.exit(2)

# Iterate through all IMAP folders and backup messages
r, response = client.list()
if 'OK' == r:
    for line in response:
        flags, delimiter, mailbox_name = list_response_pattern.match(line).groups()
        try:
            backup_imap_folder(mailbox_name)
        except Exception as e:
            logger.exception(e)
            pass
