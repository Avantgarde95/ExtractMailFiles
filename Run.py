import datetime
import json
import zipfile
import email
import base64
import os

import eml_parser


def toJSON(obj):
    if isinstance(obj, datetime.datetime):
        serial = obj.isoformat()
        return serial
    elif isinstance(obj, email.header.Header):
        print(str(obj))
        raise Exception('object cannot be of type email.header.Header')
    elif isinstance(obj, bytes):
        return obj.decode('utf-8', errors='ignore')

    raise TypeError(f'Type "{str(type(obj))}" not serializable')


def getFileFromMail(mailsZipFile, mailName):
    with mailsZipFile.open(mailName) as mailFile:
        mail = mailFile.read()
        print(f'Read {mailName}! (Size: {len(mail)})')

    return mail


def parseMail(mail):
    parser = eml_parser.EmlParser(
        include_attachment_data=True, parse_attachments=True)
    return parser.decode_email_bytes(mail)


def getAttachmentsFromEml(eml):
    if not ('attachment' in eml):
        print('-- No attachments!')
        return []

    attachments = eml['attachment']
    print(f'-- {len(attachments)} attachments!')
    return attachments


def saveAttachment(prefix, attachment):
    name = attachment['filename']
    print(f'-- Writing {name}...')
    raw = attachment['raw']
    decodedRaw = base64.b64decode(raw)

    with open(f'{prefix}{name}', 'wb') as attachmentFile:
        attachmentFile.write(decodedRaw)


def createDirectory(name):
    if not os.path.exists(name):
        os.makedirs(name)


mailsZipPath = input('Mails .zip path: ')
outputPath = input('Output directory: ')

createDirectory(outputPath)

with zipfile.ZipFile(mailsZipPath, 'r') as mailsZipFile:
    mailNames = [name for name in mailsZipFile.namelist()
                 if name.endswith('.eml')]
    print(f'Got {len(mailNames)} mails!')

    for index, mailName in enumerate(mailNames):
        mail = getFileFromMail(mailsZipFile, mailName)
        eml = parseMail(mail)
        attachments = getAttachmentsFromEml(eml)

        for attachment in attachments:
            saveAttachment(f'{outputPath}/{index + 1}_', attachment)

print('Done!')
