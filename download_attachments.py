# -*- coding: utf-8 -*-
# https://qiita.com/stkdev/items/a44976fb81ae90a66381
# https://qiita.com/zarchis/items/3258562ebc9570fa05a3
# https://torina.top/detail/119/
# --------------------------------------------------------------------
import sys
import imaplib
import email
import os
import dateutil.parser
from tqdm import tqdm
import configparser
# --------------------------------------------------------------------


def check_decode(rawemail):
    lookup = ('utf_8', 'iso2022jp', 'iso2022_jp_1', 'iso2022_jp_2', 'iso2022_jp_3',
              'shift_jis', 'shift_jis_2004', 'shift_jisx0213',
              'euc_jp', 'euc_jis_2004', 'euc_jisx0213',
              'iso2022_jp_ext', 'latin_1', 'ascii')
    for encode in lookup:
        try:
            message = email.message_from_string(rawemail.decode(encode))
            return message, encode
        except:
            pass


sys.stderr.write("*** START ***\n")

# count
cnt_already_exists = 0
cnt_downloaded = 0
cnt_none_encoding = 0
cnt_filename_uncorrect = 0
cnt_filename_none_ext = 0
cnt_filename_include_question = 0
none_extension = []
question_mark = []

# Mail Setting
config = configparser.ConfigParser()
config.read('config.ini')
config_mail = config['MAIL']

# Create Directory
detach_dir = '.'
if 'attachments' not in os.listdir(detach_dir):
    os.mkdir('attachments')

try:
    mail = imaplib.IMAP4_SSL(config_mail.get('SERVER'))
    mail.login(config_mail.get('USER'), config_mail.get('PASSWORD'))
    mail.select('Inbox')
    typ, data = mail.search(None, 'ALL')

    pbar = tqdm(total=len(data[0].split()))
    cnt = 0
    array_size = len(data[0].split())

    f = open('LastUpdatedDate.txt', 'r')
    last_updated_date = f.readline()
    f.close

    first_record_flag = True
    for num in reversed(data[0].split()):
        cnt += 1
        # Get Message
        result, d = mail.fetch(num, "(RFC822)")
        msg, msg_encoding = check_decode(d[0][1])
        if not msg_encoding:
            pbar.update(1)
            cnt_none_encoding += 1
            continue

        # Date
        date = dateutil.parser.parse(msg.get('Date')).strftime("%Y/%m/%d %H:%M:%S")
        if date <= last_updated_date:
            # Already saved attachments
            break

        if first_record_flag:
            first_record_flag = False
            f = open('LastUpdatedDate.txt', 'w')
            f.write(date)
            f.close()

        # From
        # fromObj = email.header.decode_header(msg.get('From'))
        # addr = ""
        # for f in fromObj:
        #     if isinstance(f[0], bytes):
        #         addr += f[0].decode(msg_encoding)
        #     else:
        #         addr += f[0]
        # print('From: \t\t' + addr)

        # Subject
        # subject = email.header.decode_header(msg.get('Subject'))
        # title = ""
        # for sub in subject:
        #     if isinstance(sub[0], bytes):
        #         title += sub[0].decode(msg_encoding)
        #     else:
        #         title += sub[0]
        # print('Subject: \t\t' + title)

        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            # Check File Name
            fileName = part.get_filename()
            if not fileName:
                continue

            fileName = fileName.replace('\r', '')
            fileName = fileName.replace('\n', '')
            fileName = fileName.replace('=?utf-8?Q?', '')
            fileName = fileName.replace('?=', '')
            root, ext = os.path.splitext(fileName)
            if ext is None or not ext:
                cnt_filename_none_ext += 1
                none_extension.append(fileName)
                continue
            elif ext.find('?') > -1:
                cnt_filename_include_question += 1
                question_mark.append(fileName)
                # continue

            if bool(fileName):
                filePath = os.path.join(detach_dir, 'attachments', fileName)
                if not os.path.isfile(filePath):
                    fp = open(filePath, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
                    cnt_downloaded += 1
                else:
                    cnt_already_exists += 1
        pbar.update(1)
except Exception as ee:
    sys.stderr.write("*** Error Occured ***\n")
    sys.stderr.write(str(ee) + '\n')
finally:
    if mail is not None:
        mail.close()
        mail.logout()
    if pbar is not None:
        pbar.close()
    sys.stderr.write("*** End ***\n")

print('\nDownloaded: \t{}'.format(cnt_downloaded))
print('Already Exists: {}'.format(cnt_already_exists))
print('Encoding Error: {}'.format(cnt_none_encoding))
print('None Extension: {}'.format(cnt_filename_none_ext) + ' ({})'.format(none_extension))
print('Include ? Mark: {}'.format(cnt_filename_include_question) + ' ({})'.format(question_mark))
