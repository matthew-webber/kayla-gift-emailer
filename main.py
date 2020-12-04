from string import Template
import os
import pathlib
import csv
import sys


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


# todo move to custom
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
        files = os.listdir()
    else:
        files = os.listdir(dir)

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


def main(**kwargs):

    mail_mode = kwargs.get('mail_mode')  # tells the mail sender whether to display/save/send
    iter_start_row = kwargs.get('row_number')  # tells the mail sender whether to display/save/send
    records_per_loop = kwargs.get('iteration_number')  # tells the mail sender whether to display/save/send

    mail_mode = mail_mode if mail_mode else 'Display'
    iter_start_row = int(iter_start_row) if iter_start_row else 2

    records_per_loop = int(records_per_loop) if records_per_loop else 3

    # establish globals
    EMAIL_TALLY = 0
    GIVER_MAIL_SUBJECT = 'Thank You for Your Aquarium Gift Membership Purchase!'
    RECIPIENT_MAIL_SUBJECT = 'You\'ve Been Given the Gift of Membership to the South Carolina Aquarium!'

    # column numbers as of 11/24
    GIVER_FULLNAME_COL = 56  # e.g. "Susan Blender" $gift_giver_fullname
    GIVER_FIRSTNAME_COL = 11
    GIVER_EMAIL_COL = 9
    GIVER_NICKNAME_COL = 38

    RECIPIENT_FULLNAME_COL = 37
    RECIPIENT_FIRSTNAME_COL = 59  # e.g. "Erica" $recipient_first_name refactor firstname_col above
    RECIPIENT_EMAIL_COL = 61

    MEMLEVEL_COL = 54
    EXPIRATION_COL = 51
    MESSAGE_COl = 50

    # establish helper objects
    working_row_set = []  # contains dicts of rows for which emails are currently being generated
    reader_storage = []  # takes rows from csv.reader so file can close / # rows determined / etc.
    im_done = False  # todo refactor terminate program flag name

    # for Mac
    if os.name == 'posix':
        csv_file = find_first_with_ext_in_dir('csv')
        giver_template = './templates/giver_template.html'
        recipient_template = './templates/recipient_template.html'
    elif os.name == 'nt':
        csv_file = find_first_with_ext_in_dir('csv')  # todo refactor when in production
        giver_template = f'{get_pwd_of_this_file()}\\templates\\giver_template.html'
        recipient_template = f'{get_pwd_of_this_file()}\\templates\\recipient_template.html'

    with open(csv_file, 'r') as f:
        member_reader = csv.reader(f)

        for row in member_reader:
            reader_storage.append(row)

    run_loop = True

    while run_loop is True:

        if len(reader_storage) > iter_start_row + records_per_loop:
            iter_end_row = iter_start_row + records_per_loop
        else:
            iter_end_row = iter_start_row + (len(reader_storage) - iter_start_row)
            im_done = True

        # '- 1' to accommodate for 0-index
        for row in reader_storage[iter_start_row - 1:iter_end_row]:
            working_row_set.append(dict(
                gift_giver_fullname=row[GIVER_FULLNAME_COL - 1],
                giver_first_name=row[GIVER_FIRSTNAME_COL - 1],
                giver_identification=row[GIVER_NICKNAME_COL - 1],
                giver_email=row[GIVER_EMAIL_COL - 1],
                recipient_full_name=row[RECIPIENT_FULLNAME_COL - 1],
                recipient_first_name=row[RECIPIENT_FIRSTNAME_COL - 1],
                recipient_email=row[RECIPIENT_EMAIL_COL - 1],
                gift_message=f'<em>"{row[MESSAGE_COl - 1]}"</em>',
                membership_expiration=row[EXPIRATION_COL - 1],
                membership_level=row[MEMLEVEL_COL - 1],
                ))

        total_records = len(reader_storage) - 1  # -1 for header row
        records_remaining = len(reader_storage) - iter_end_row
        print(f'Processing {len(working_row_set)} records...')

        # for each record, generate an email
        for record in working_row_set:

            if os.name == 'posix':
                # generate giver email
                with open(giver_template, 'r') as f:
                    t = Template(f.read())

                posix_run(mail_subject=GIVER_MAIL_SUBJECT,
                          recipients=[record['giver_email']],
                          template_vals=record,
                          template=t,
                          tally=EMAIL_TALLY,
                          )

                EMAIL_TALLY += 1

                # generate recipient email
                with open(recipient_template, 'r') as f:
                    t = Template(f.read())

                posix_run(mail_subject=RECIPIENT_MAIL_SUBJECT,
                          recipients=[record['recipient_email']],
                          template_vals=record,
                          template=t,
                          tally=EMAIL_TALLY,
                          )

                EMAIL_TALLY += 1

                print('\n\n\n')

            elif os.name == 'nt':

                # generate giver email
                with open(giver_template, 'r') as f:
                    t = Template(f.read())

                    send_outlook_html_mail(recipients=[record['giver_email']], subject=GIVER_MAIL_SUBJECT, body=t.substitute(record),
                                           message_action=mail_mode)

                # generate recipient email
                with open(recipient_template, 'r') as f:
                    t = Template(f.read())

                send_outlook_html_mail(recipients=[record['recipient_email']], subject=RECIPIENT_MAIL_SUBJECT, body=t.substitute(record),
                                       message_action=mail_mode)

        print('...Done!')
        print(f'Total emails generated: {EMAIL_TALLY}')

        if iter_end_row == total_records:
            print('Script has reached the end of the file!')
            run_loop = False  # end program

        else:

            print(f'Records remaining: {records_remaining}')
            continue_loop = True

            print('Press enter to continue or q to quit.')

            while continue_loop is True:

                x = input('?:').strip().lower()

                if x == '':

                    continue_loop = False
                    iter_start_row = iter_start_row + records_per_loop
                    working_row_set = []

                elif x == 'q':

                    continue_loop = False
                    run_loop = False  # end program

                else:

                    print('Unknown response.  Try again.')
                    continue


if __name__ == '__main__':

    DEFAULT_ROW_NUMBER = 2
    DEFAULT_ITER_NUMBER = 3

    try:
        row_number = sys.argv[1]
    except IndexError:
        row_number = DEFAULT_ROW_NUMBER

    try:
        iteration_number = sys.argv[2]
    except IndexError:
        iteration_number = DEFAULT_ITER_NUMBER

    # start the CLI
    try:
        with open('templates/prompt.txt', 'r') as f:
            t = Template(f.read())
    except FileNotFoundError:
        with open(f'{get_pwd_of_this_file()}\\templates\\prompt.txt', 'r') as f:
            t = Template(f.read())

    prompt = t.substitute(dict(row_number=row_number, iteration_number=iteration_number))

    cli_selectors = dict(
        start=['start'],
        display=['display'],
        quit=['quit', 'q', 'exit'],
    )

    print(prompt)

    while True:

        resp = input("?:").strip().lower()

        if resp in cli_selectors.get('start'):

            mail_mode = None

        elif resp in cli_selectors.get('display'):

            mail_mode = 'Display'

        elif resp in cli_selectors.get('quit'):

            pass

        else:

            print('Unknown response.  Try again.')
            continue

        break

    if resp != 'quit':
        main(mail_mode=mail_mode, row_number=row_number, iteration_number=iteration_number)

    say_goodbye()
