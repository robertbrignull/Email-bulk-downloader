#!/usr/bin/python

import os
import sys
import imaplib
import getpass
import email
import email.header

def make_filename_safe(filename):
    return filename.replace('\0', '').replace('/', '_')

host = raw_input("Host: ")
username = raw_input("Email address: ")
password = getpass.getpass()

M = imaplib.IMAP4_SSL(host)

try:
    rv, data = M.login(username, password)
except imaplib.IMAP4.error:
    print "LOGIN FAILED!!! "
    sys.exit(1)

rv, mailboxes = M.list()
if rv == 'OK':
    print "Mailboxes: {0}".format(map(lambda (x): x.split(' "." ')[1][1:-1], mailboxes))

if len(mailboxes) > 0:
    mailbox = raw_input("Mailbox to download: ")
    output_dir = raw_input("Output directory: ")

    output_dir = os.path.join(os.getcwd(), output_dir, username)

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    rv, data = M.select(mailbox)
    if rv == 'OK':
        print "Processing mailbox..."
        
        rv, data = M.search(None, "ALL")
        if rv == 'OK':
            print "{0} messages found".format(len(data[0].split()))
        else:
            print "No messages found!"

        for num in data[0].split():
            rv, data = M.fetch(num, '(RFC822)')
            if rv == 'OK':
                print "Processing message {0}".format(num)

                msg = email.message_from_string(data[0][1])
                date = email.utils.parsedate_tz(msg['Date'])
                filename_prefix = num + '_' + str(date[0]) + '-' + str(date[1]).zfill(2) + '-' + str(date[2]).zfill(2)
                filename = make_filename_safe(filename_prefix + '_' + msg['Subject'] + '.eml')
                out_file = open(os.path.join(output_dir, filename), 'w')
                out_file.write("From: {0}\n".format(msg['From']))
                out_file.write("To: {0}\n".format(msg['To']))
                out_file.write("Subject: {0}\n".format(msg['Subject']))

                for part in msg.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue

                    if part.get_content_type() == 'text/plain':
                        out_file.write("\n\n")
                        out_file.write(part.as_string())

                    if part.get_filename() != None:
                        filePath = os.path.join(output_dir, filename_prefix + '_' + make_filename_safe(part.get_filename()))
                        attachment_file = open(filePath, 'wb')
                        attachment_file.write(part.get_payload(decode=True))
                        attachment_file.close()

                out_file.close()
            else:
                print "ERROR getting message {0}".format(num)

        M.close()
    else:
        print "ERROR: Unable to open mailbox {0}".format(rv)

M.logout()
