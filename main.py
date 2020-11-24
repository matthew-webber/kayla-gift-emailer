from string import Template
# import win32com.client

# Hard coded email subject
MAIL_SUBJECT = 'AUTOMATED Text Python Email without attachments'

# Hard coded email HTML text
MAIL_BODY = \
    '<html> ' \
    ' <body>' \
    ' <p><b>Dear</b> Receipient,<br><br>' \
    ' This is an automatically generated email by <font size="5" color="blue">Python.</font><br>' \
    ' It is so <del>amazing and</del> fantastic<br>' \
    ' <strong>Wish you</strong> all the <font size="5" color="green">best</font><br>' \
    ' </body>' \
    '</html>'


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
    # if len(recipients) > 0 and isinstance(recipient_list, list):
    #     outlook = win32com.client.Dispatch("Outlook.Application")
    #
    #     ol_msg = outlook.CreateItem(0)
    #
    #     str_to = ""
    #     for recipient in recipients:
    #         str_to += recipient + ";"
    #
    #     ol_msg.To = str_to
    #
    #     if copies is not None:
    #         str_cc = ""
    #         for cc in copies:
    #             str_cc += cc + ";"
    #
    #         ol_msg.CC = str_cc
    #
    #     ol_msg.Subject = subject
    #     ol_msg.HTMLBody = body
    #
    #     if send_or_display.upper() == 'SEND':
    #         ol_msg.Send()
    #     else:
    #         ol_msg.Display()
    # else:
    #     print('Recipient email address - NOT FOUND')


if __name__ == '__main__':

    a = dict(
        giver_first_name='John',
        membership_level = 'Family',
        recipient_full_name = 'Darla May',
        giver_identification = 'Ur luver boy John Boy',
        gift_message = 'Luv u honey dumplin',
        membership_expiration = '1/1/2099'
    )

    from string import Template

    with open("./template.html", 'r') as f:
        t = Template(f.read())
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

    # recipient_list = ['mattwebbersemail@gmail.com', 'jake@example.com']
    #
    # copies_list = ['mattwebbersemail@gmail.com', 'jake@example.com']
    #
    # for i in range(len(recipient_list)):
    #
    #     send_outlook_html_mail(recipients=[recipient_list.pop()], subject=MAIL_SUBJECT, body=MAIL_BODY,
    #                            send_or_display='Display',
    #                            copies=[copies_list.pop()])
    #
    #     x = input('Continue?')
    #
    #     if x == 'y':
    #
    #         continue
    #
    #     else:
    #
    #         break

