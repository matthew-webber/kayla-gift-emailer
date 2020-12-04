from string import Template
import os
import pathlib
import csv
import sys
from funcs import *
from query_template_matcher import RecordData


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

    # CSV COLUMN NUMBERS (as of 11/24)
    QUERY_NAME_COL = 1

    GIVER_FULLNAME_COL = 56  # e.g. "Susan Blender" $giver_fullname
    SALUTATION = 11
    GIVER_EMAIL_COL = 9
    GIVER_NICKNAME_COL = 38

    RECIPIENT_FULLNAME_COL = 37
    RECIPIENT_FIRSTNAME_COL = 59  # e.g. "Erica" $recipient_first_name-- giver equivalent is "salutation"
    RECIPIENT_EMAIL_COL = 61

    GUARDIAN_FIRSTNAME = 62  # col BJ - will be different than recipient_firstname so added to end
    GUARDIAN_STG_ORDERNOTES = 52  # col AZ

    MEMLEVEL_COL = 54
    EXPIRATION_COL = 51
    MESSAGE_COl = 50

    # establish helper objects
    working_row_set = []  # contains dicts of rows for which emails are currently being generated
    reader_storage = []  # takes rows from csv.reader so file can close / # rows determined / etc.
    im_done = False  # todo confirm can remove

    data_object = RecordData()
    # print(data_object.csv_file + ' is do csv file')
    # print(data_object.queries + ' is do csv file')

    with open(data_object.csv_file, 'r') as f:
        member_reader = csv.reader(f)

        for row in member_reader:
            reader_storage.append(row)

    run_loop = True

    # begin iteration over each record
    while run_loop is True:

        if len(reader_storage) > iter_start_row + records_per_loop:
            iter_end_row = iter_start_row + records_per_loop
        else:
            iter_end_row = iter_start_row + (len(reader_storage) - iter_start_row)
            im_done = True

        # '- 1' to accommodate for 0-index
        for row in reader_storage[iter_start_row - 1:iter_end_row - 1]:

            # adjust the gift message value if none included
            if row[MESSAGE_COl -1] == '':
                row[MESSAGE_COl -1] = 'Enjoy your membership!'
            print(row[GIVER_FULLNAME_COL - 1])
            working_row_set.append(dict(
                giver_fullname=row[GIVER_FULLNAME_COL - 1],
                salutation=row[SALUTATION - 1],
                giver_identification=row[GIVER_NICKNAME_COL - 1],
                emails=[row[GIVER_EMAIL_COL - 1], row[RECIPIENT_EMAIL_COL - 1]],
                recipient_full_name=row[RECIPIENT_FULLNAME_COL - 1],
                recipient_first_name=row[RECIPIENT_FIRSTNAME_COL - 1],
                gift_message=f'<em>"{row[MESSAGE_COl - 1]}"</em>',
                membership_expiration=row[EXPIRATION_COL - 1],
                membership_level=row[MEMLEVEL_COL - 1],
                stg_online_order_notes_1=row[GUARDIAN_STG_ORDERNOTES - 1],
                guardian_first_name=row[GUARDIAN_FIRSTNAME - 1],
                query_name=row[QUERY_NAME_COL - 1],
                ))

        total_records = len(reader_storage) - 1  # -1 for header row
        records_remaining = len(reader_storage) - iter_end_row
        print(f'Processing {len(working_row_set)} records...')

        # for each record, generate an email
        for record in working_row_set:
            # set new data_obj.templates / data_obj.subjects
            data_object.reset_record_data(record['query_name'])

            # # get template filestring
            # template_files = get_email_template(QUERY_NAME_COL - 1, row, template_dict)

            # todo the amount is never not 2 at this point so why refactor?
            for i in range(2):

                if os.name == 'posix':

                    with open(data_object.templates[i], 'r') as f:
                        t = Template(f.read())

                    posix_run(mail_subject=data_object.subjects[i],
                              recipients=[record['emails'][i]],
                              template_vals=record,
                              template=t,
                              tally=EMAIL_TALLY,
                              )

                    EMAIL_TALLY += 1

                    print('\n\n\n')

                elif os.name == 'nt':

                    # refactor like generate_emails([template1, template2])
                    # generate giver email
                    with open(data_object.templates[i], 'r') as f:
                        t = Template(f.read())

                    send_outlook_html_mail(recipients=[record['emails'][i]],
                                           subject=data_object.subjects[i],
                                           body=t.substitute(record),
                                           message_action=mail_mode
                                           )

                    EMAIL_TALLY += 1


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

            mail_mode = 'Save'

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
