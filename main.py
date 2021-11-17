from string import Template
import os
import csv
import sys
from funcs import *
from query_template_matcher import RecordData

def main(**kwargs):

    # tells the mail sender whether to display/save/send
    mail_action = kwargs.get('mail_action')
    iter_start_row = int(kwargs.get('row_number'))
    records_per_loop = int(kwargs.get('iteration_number'))
    data = kwargs.get('data')

    mail_action = mail_action if mail_action else 'Display'

    # convert column letters to numbers from JSON data
    for column in data['columns'].keys():
        col_letters = data['columns'][column]
        data['columns'][column] = excel_col_to_number(col_letters)

    # track number emails/rows processed
    email_tally = 0
    record_tally = 0

    MESSAGE_COl = 50

    # establish helper objects
    
    # contains dicts of rows for which emails are currently being generated
    working_row_set = []
    # takes rows from csv.reader so file can close / # rows determined / etc.
    reader_storage = []

    data_object = RecordData(data)

    with open(data_object.csv_file, 'r') as f:
        member_reader = csv.reader(f)

        for row in member_reader:
            reader_storage.append(row)

    run_loop = True

    # begin iteration over each record
    while run_loop is True:

        if len(reader_storage) >= iter_start_row + records_per_loop:
            iter_end_row = iter_start_row + records_per_loop
            im_done = False
        else:
            # because of the way I'm using the "start/end rows" to slice the reader_storage variable, using the len() of
            # reader_storage and subtracting the "human readable" iter_start_row results in a end row that is off by 1
            iter_end_row = iter_start_row + \
                (len(reader_storage) - iter_start_row) + \
                1  # todo math this better
            im_done = True

        # start processing of row "chunk" before pausing for user input to continue or quit
        # '- 1' to accommodate for 0-index
        for row in reader_storage[iter_start_row - 1:iter_end_row - 1]:

            record_tally += 1

            # set the template values for the row being processed
            if row[MESSAGE_COl - 1] == '':
                # adjust the gift message value if none included
                row[MESSAGE_COl - 1] = 'Enjoy your membership!'
            working_row_set.append(dict(
                giver_fullname=row[data['columns']['giverFullName'] - 1],
                salutation=row[data['columns']['giverSalutation'] - 1],
                giver_identification=row[data['columns']['giverNickname'] - 1],
                emails=[row[data['columns']['giverEmail'] - 1],
                        row[data['columns']['recipientEmail'] - 1]],
                recipient_full_name=row[data['columns']
                                        ['recipientFullName'] - 1],
                recipient_first_name=row[data['columns']
                                         ['recipientFirstName'] - 1],
                gift_message=f'<em>"{row[data["columns"]["giftMessage"] - 1]}"</em>',
                membership_expiration=row[data['columns']
                                          ['membershipExpiration'] - 1],
                membership_level=row[data['columns']['membershipLevel'] - 1],
                stg_online_order_notes_1=row[data['columns']
                                             ['guardianOrderNotes'] - 1],
                guardian_first_name=row[data['columns']
                                        ['guardianFirstName'] - 1],
                query_name=row[data['columns']['queryName'] - 1],
            ))

        # total_records = len(reader_storage) - 1  # -1 for header row
        records_remaining = len(reader_storage) - record_tally - 1
        print(f'Processing {len(working_row_set)} records...')

        # for each record, generate an email
        for record in working_row_set:

            # set new data_obj.templates / data_obj.subjects
            data_object.reset_record_data(record['query_name'])

            for i in range(2):

                if os.name == 'posix':

                    with open(data_object.templates[i], 'r') as f:
                        t = Template(f.read())

                    posix_run(mail_subject=data_object.subjects[i],
                              recipients=[record['emails'][i]],
                              template_vals=record,
                              template=t,
                              tally=email_tally,
                              )

                    email_tally += 1

                    print('\n\n\n')

                elif os.name == 'nt':

                    # refactor like generate_emails([template1, template2])
                    # generate giver email
                    with open(data_object.templates[i], 'r') as f:
                        t = Template(f.read())

                    send_outlook_html_mail(recipients=[record['emails'][i]],
                                           subject=data_object.subjects[i],
                                           body=t.substitute(record),
                                           message_action=mail_action
                                           )

                    email_tally += 1

        print('...Done!')
        print(f'Total emails generated: {email_tally}')
        print(
            f'Records processed: {record_tally} of {len(reader_storage) - 1}')

        if im_done:
            print('All records processed!')
            run_loop = False  # end program

        else:
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

    from json import load as json_load

    # todo refactor to posix vs nt
    try:
        with open('project.json', 'r') as f:
            project_data = json_load(f)
    except FileNotFoundError:
        with open(f'{get_pwd_of_this_file()}\\project.json', 'r') as f:
            project_data = json_load(f)

    defaults = project_data['default']

    try:
        row_number = sys.argv[1]
    except IndexError:
        row_number = defaults['startRow']

    try:
        record_batch = sys.argv[2]
    except IndexError:
        record_batch = defaults['recordBatch']

    # start the CLI
    # todo refactor to posix vs nt
    try:
        with open('cli-prompt.txt', 'r') as f:
            t = Template(f.read())
    except FileNotFoundError:
        with open(f'{get_pwd_of_this_file()}\\cli-prompt.txt', 'r') as f:
            t = Template(f.read())

    prompt = t.substitute(
        dict(row_number=row_number, iteration_number=record_batch))

    cli_selectors = dict(
        start=['start', ''],
        display=['display'],
        quit=['quit', 'q', 'exit'],
        send=['send'],
    )

    print(prompt)

    while True:

        resp = input("?:").strip().lower()
        # todo comment out after dev over
        # todo add an env var checker for "prod" vs "dev"
        # resp = 'start'

        # check that the user has added the column for recipient email
        # before beginning and allow override
        passed, msg = check_recipient_email_present()
        print(msg)
        if not passed:
            resp = 'quit'

        if resp in cli_selectors.get('start'):

            action = 'Save'

        elif resp in cli_selectors.get('display'):

            action = 'Display'

        elif resp in cli_selectors.get('send'):

            action = 'Send'

        elif resp in cli_selectors.get('quit'):

            pass

        else:

            print('Unknown response.  Try again.')
            continue

        break

    if resp not in cli_selectors.get('quit'):
        print('\n\n......STARTING......\n\n')
        main(mail_action=action,
             row_number=row_number,
             iteration_number=record_batch,
             data=project_data)

    say_goodbye()


# todo add a better summary of what was generated instead of just one name per record
# todo add a more helpful readme for reminding how to run the thing, how to update the json file, etc.
# todo move the member_data.csv to an examples folder and say "if no csv in dir, get the example one"
# todo add the "custom" thing to the project instead of being a submodule / something that needs to be dl'ed
# todo add .bat files for easy running on Windows
