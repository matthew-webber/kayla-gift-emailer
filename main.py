from string import Template
import os
import pathlib
import win32com.client

# Hard coded email subject
MAIL_SUBJECT = 'Thank You for Your Aquarium Gift Membership Purchase!'

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


if __name__ == '__main__':

    import csv

    # establish globals
    CSV_ROW_BATCH_SIZE = 5  # the number of rows to take from the CSV file for each iteration
    START_ROW = 1  # assumes headers on row 1 and data starts on row 2 but w/ zero-index -- therefore: 1

    # column numbers as of 11/24
    FIRSTNAME_COL = 11
    RECIPIENT_COL = 37
    GIVER_COL = 38
    MESSAGE_COl = 50
    EXPIRATION_COL = 51
    MEMLEVEL_COL = 54

    # establish helper objects
    working_row_set = []  # contains dicts of rows for which emails are currently being generated
    reader_storage = []
    im_done = False

    # get csv file full path
    csv_file = find_first_with_ext_in_dir('csv')

    # with open(csv_file, 'r') as f:
    with open('C:\\Users\\Matthew Webber\\Desktop\\kayla-gift-emailer-master\\member_data.csv', 'r') as f:
        member_reader = csv.reader(f)

        for row in member_reader:
            reader_storage.append(row)

    while True:

        if len(reader_storage) > START_ROW + CSV_ROW_BATCH_SIZE:
            END_ROW = START_ROW + CSV_ROW_BATCH_SIZE
        else:
            END_ROW = START_ROW + (len(reader_storage) - START_ROW)
            im_done = True

        for row in reader_storage[START_ROW:END_ROW]:
            working_row_set.append(dict(
                giver_first_name=row[FIRSTNAME_COL - 1],
                recipient_full_name=row[RECIPIENT_COL - 1],
                giver_identification=row[GIVER_COL - 1],
                gift_message=row[MESSAGE_COl - 1],
                membership_expiration=row[EXPIRATION_COL - 1],
                membership_level=row[MEMLEVEL_COL - 1],
            ))

        # load the template up
        from string import Template

        # with open("./giver_template.html", 'r') as f:
        with open("C:\\Users\\Matthew Webber\\Desktop\\kayla-gift-emailer-master\\giver_template.html") as f:
            t = Template(f.read())

        # for each record, generate an email
        for _ in working_row_set:
            # todo add email column to the above and remove the testing crap from the kwarg situation below
            send_outlook_html_mail(recipients=['jake@example.com'], subject=MAIL_SUBJECT, body=t.substitute(_),
                                   send_or_display='Display',
                                   copies=['jake@example.com'])

        if im_done == True:
            break
        else:

            x = input('Continue?')

            if x == 'y':

                continue

            else:

                break

            START_ROW = START_ROW + CSV_ROW_BATCH_SIZE
            working_row_set = []
            continue


    # JUST SOME PSEUDO-CODE BELOW
    # ////////////////////////////
    # ////////////////////////////
    # ////////////////////////////
    # ////////////////////////////
    # from string import Template
    #
    # # with open("./giver_template.html", 'r') as f:
    # with open("C:\\Users\\Matthew Webber\\Desktop\\kayla-gift-emailer-master\\giver_template.html") as f:
    #     t = Template(f.read())

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
