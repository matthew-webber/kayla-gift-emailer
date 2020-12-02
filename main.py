from string import Template
import os
import pathlib
import csv
import sys


def say_goodbye():
    print('Terminating....')
    print('\n\n~~~~ Goodbye ~~~~~')


def posix_run(mail_subject, recipients, template, template_vals):
    print('////////////////////////////\n///////////////////////////////////')
    print('NEW MESSAGE')
    print('recipients : ' + str(recipients))
    print('subject : ' + mail_subject)
    print('body : \n\n' + template.substitute(template_vals) + '\n\n')

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
            return f'{get_pwd_of_this_file()}/{file}'


def send_outlook_html_mail(recipients, subject='No Subject', body='Blank', send_or_display='Display', copies=None):
    """
    Send an Outlook HTML email
    :param recipients: list of recipients' email addresses (list object)
    :param subject: subject of the email
    :param body: HTML body of the email
    :param send_or_display: Send - send email automatically | Display - email gets created user have to click Send
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

        if send_or_display.upper() == 'SEND':
            ol_msg.Send()
        else:
            ol_msg.Display()
    else:
        print('Recipient email address - NOT FOUND')


def main():

    # establish globals
    CSV_ROW_BATCH_SIZE = 3  # the number of rows to take from the CSV file for each iteration
    START_ROW = 1  # assumes headers on row 1 and data starts on row 2 but w/ zero-index -- therefore: 1
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
    y_n_selectors = dict(  # determine control flow from user input
        y=['yes', 'y', 'ok', ':)'],
        n=['no', 'n', ':(']
    )

    # for Mac
    if os.name == 'posix':
        csv_file = find_first_with_ext_in_dir('csv')
        giver_template = './templates/giver_template.html'
        recipient_template = './templates/recipient_template.html'
    elif os.name == 'nt':
        csv_file = 'C:\\Users\\Matthew Webber\\Desktop\\kayla-gift-emailer-master\\member_data.csv'  # todo refactor when in production
        # csv_file = find_first_with_ext_in_dir('csv')
        giver_template = 'C:\\Users\\Matthew Webber\\Desktop\\kayla-gift-emailer-master\\templates\\giver_template.html'
        recipient_template = 'C:\\Users\\Matthew Webber\\Desktop\\kayla-gift-emailer-master\\templates\\recipient_template.html'

    with open(csv_file, 'r') as f:
        member_reader = csv.reader(f)

        for row in member_reader:
            reader_storage.append(row)

    run_loop = True

    while run_loop is True:

        if len(reader_storage) > START_ROW + CSV_ROW_BATCH_SIZE:
            END_ROW = START_ROW + CSV_ROW_BATCH_SIZE
        else:
            END_ROW = START_ROW + (len(reader_storage) - START_ROW)
            im_done = True

        for row in reader_storage[START_ROW:END_ROW]:
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

        print(f'Attempting to generate {len(working_row_set)}/{len(reader_storage)} emails...')

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
                          )

                # generate recipient email
                with open(recipient_template, 'r') as f:
                    t = Template(f.read())
                posix_run(mail_subject=RECIPIENT_MAIL_SUBJECT,
                          recipients=[record['recipient_email']],
                          template_vals=record,
                          template=t,
                          )

            elif os.name == 'nt':

                # generate giver email
                with open(giver_template, 'r') as f:
                    t = Template(f.read())

                    send_outlook_html_mail(recipients=[record['giver_email']], subject=GIVER_MAIL_SUBJECT, body=t.substitute(record),
                                           send_or_display='Display')

                # generate recipient email
                with open(recipient_template, 'r') as f:
                    t = Template(f.read())

                send_outlook_html_mail(recipients=[record['recipient_email']], subject=RECIPIENT_MAIL_SUBJECT, body=t.substitute(record),
                                       send_or_display='Display')

        print('...Done!')

        if im_done is True:
            print('Script has reached the end of the file!')
            say_goodbye()
            run_loop = False  # end program

        else:

            continue_loop = True

            while continue_loop is True:

                x = input('Continue? [y/n]  ?: ')

                if x in y_n_selectors.get('y'):

                    continue_loop = False
                    START_ROW = START_ROW + CSV_ROW_BATCH_SIZE
                    working_row_set = []

                elif x in y_n_selectors.get('n'):

                    say_goodbye()
                    continue_loop = False
                    run_loop = False  # end program

                else:

                    print('Response not valid. Try again.')


if __name__ == '__main__':

    DEFAULT_ROW_NUMBER = 2
    DEFAULT_ITER_NUMBER = 3

    try:
        row_number = sys.argv[1]
        print(0)
    except IndexError:
        row_number = DEFAULT_ROW_NUMBER

    try:
        row_number = sys.argv[2]
        print(1)
    except IndexError:
        iteration_number = DEFAULT_ITER_NUMBER

    # start the CLI
    try:
        with open('templates/cli.txt', 'r') as f:
            t = Template(f.read())
    except FileNotFoundError:
        with open('C:\\Users\\Matthew Webber\\Desktop\\kayla-gift-emailer-master\\templates\\cli.txt', 'r') as f:
            t = Template(f.read())

    prompt = t.substitute(dict(row_number=row_number, iteration_number=iteration_number))

    resp = input(prompt)

    if resp.strip().lower() == 'start':
        main()
    else:
        print('\n\nGoodbye!')



    # JUST SOME PSEUDO-CODE BELOW
    # ///////////////////////////////////////////////////////////////////////////////
    # ///////////////////////////////////////////////////////////////////////////////
    # /////////////////////////////////////////////////////////////////////////////////
    # //////////////////////////////////////////////////////////////////////////////

    # get .csv file in current directory
    # load up to reader obj
    # while True
    #   for every USER_DEFINED_NUMBER rows:
    #     load variables for sets 1 + 2 according to whatever columns are needed
    #     call "send_outlook_html_mail" to variable set 1
    #     call "send_outlook_html_mail" to variable set 2
    #     keeping running total, inform user how many rows have been processed so far
    #
    #   ask user if they want to continue
    #   if yes, process next USER_DEFINED_NUMBER rows in same spreadsheet
    #   else, break
    # end program
