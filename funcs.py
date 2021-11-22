from string import ascii_lowercase
import os
import pathlib

def say_goodbye():
    print('\nTerminating....')
    print('~~~~ Goodbye ~~~~~')


def posix_run(mail_subject, recipients, template, template_vals, tally):
    print('////////////////////////////\n////////////////////////////')
    print('NEW MESSAGE')
    print('recipients : ' + str(recipients))
    print('subject : ' + mail_subject)
    print(f'email number {tally + 1}')
    # comment out for dev
    # print('body : \n\n' + template.substitute(template_vals) + '\n\n')


# todo move to custom
def get_pwd_of_this_file():
    """
    Assumes os
    Gets the pwd of the file running this thing
    :return: path string
    """
    return os.path.dirname(os.path.realpath(__file__))


# todo this doesn't really work right if you add a dir to the arguments
def find_first_with_ext_in_dir(extension, dir=None):
    """
    Assumes os, pathlib
    Returns path of first file matching a given extension in the given dir
    :param extension: e.g. 'csv', 'pdf', etc.
    :param dir: path string
    :return: path string
    """

    if dir is None:
        files = os.listdir(get_pwd_of_this_file())
    else:
        files = os.listdir(dir)
    print(os.listdir())
    # print('above is os.listdir')
    # print(dir + ' is dir')

    if extension[0] != '.':
        extension = '.' + extension

    for file in files:
        if pathlib.Path(file).suffix == extension:
            if os.name == 'posix':
                return f'{get_pwd_of_this_file()}/{file}'
            elif os.name == 'nt':
                return f'{get_pwd_of_this_file()}\\{file}'


def send_outlook_html_mail(recipients, subject='No Subject', body='Blank', message_action='Display', copies=None):
    """ 
    Send an Outlook HTML email
    :param recipients: list of recipients' email addresses (list object)
    :param subject: subject of the email
    :param body: HTML body of the email
    :param message_action: Send - send email automatically | Display - email gets created user have to click Send
    :param copies: list of CCs' email addresses
    :return: None
    """

    import win32com
    import win32com.client

    if len(recipients) > 0:
        # and isinstance(recipient_list, list) \
        outlook = win32com.client.Dispatch("Outlook.Application")

        ol_msg = outlook.CreateItem(0)

        str_to = ""
        for recipient in recipients:
            str_to += recipient + ";"

        ol_msg.To = str_to

        if copies is not None:
            str_cc = ""
            for cc in copies:
                str_cc += cc + ";"

            ol_msg.CC = str_cc

        ol_msg.Subject = subject
        ol_msg.HTMLBody = body

        if message_action.upper() == 'SEND':
            ol_msg.Send()
        elif message_action.upper() == 'SAVE':
            ol_msg.Save()
        else:
            ol_msg.Display()
    else:
        print('Recipient email address - NOT FOUND')


def get_email_template(query_column, record, template_dict):
    """
    Checks the query associated with a record and returns a list of file strings
    corresponding to the templates in this project which are associated with the
    given query in the record.
    :param column_number: index of query column
    :param record: list of all fields in the record
    :param template_dict: query-template dict
    :return: file string to template
    """

    return template_dict.get(record[query_column])

def check_recipient_email_present():
    
    from json import load as json_load
    import csv
    from query_template_matcher import RecordData
    
    print('1')
    try:
        with open('project.json', 'r') as f:
            data = json_load(f)
    except FileNotFoundError:
        with open(f'{get_pwd_of_this_file()}\\project.json', 'r') as f:
            data = json_load(f)
    print('2')

    x = RecordData(data=data)
    x.reset_record_data('MEM-Gift_Primary_Web Giver Inc_Acknowledgement Letter')
    
    with open(x.csv_file, 'r') as f:
        member_reader = csv.reader(f)
        reader_storage = list()
        
        for row in member_reader:
            reader_storage.append(row)

    recipient_email_col = data["columns"]["recipientEmail"]  # store for msg just in case (being lazy bc data is written over below to a column number)
    
    # convert column letters to numbers from JSON data
    
    for column in data['columns'].keys():
        col_letters = data['columns'][column]
        data['columns'][column] = excel_col_to_number(col_letters)
    
    msg = ''
    # print(data['columns']['recipientEmail'])
    for i, row in enumerate(reader_storage):
        msg += f'\nRow {i + 1} - '
        try:
            msg += f'{row[data["columns"]["recipientEmail"] - 1]}'
        except IndexError:
            msg += f'\n\n***WARNING*** It looks like you don\'t have a "Recipient Email" (column {recipient_email_col}) or there is a missing value at the above row.'
            return (False, msg)
    
    return (True, 'INFO: Recipient column passed test.')


def excel_col_to_number(col):
    """
    Converts a column name (in letter form) to a number

    :param col: a column name like "A", "Y", "DBZ", etc.
    :return: the number corresponding to the column
    """

    num_equivalents = list()
    answer = 0

    # "ABC" => [1, 2, 3]
    for letter in col:
        num_equivalents.append((ascii_lowercase.index(letter.lower()) + 1))

    answer = 0

    for i, p in enumerate(reversed(num_equivalents)):
        if i == 0:
            answer += p
            continue

        answer += ((26 ** i) * p)

    return answer
