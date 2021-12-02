from funcs import *
import os


class RecordData:

    if os.name == 'posix':
        sep = '/'
        pwd = '.'
    elif os.name == 'nt':
        sep = '\\'
        pwd = get_pwd_of_this_file()

    # so that I remember for later when I refactor this mess...
    # list indices 1,2 are Membership V1/A ... 3,4 are Mem V2/B ... 5,6 are STG
    # I was in a pinch and didn't have time to make a factory class for producing
    # the objects I needed so I just hard-coded the templates which matched subjects
    # using lists

    templates = [
        f'{pwd}{sep}templates{sep}giver_webgiverinc.html',
        f'{pwd}{sep}templates{sep}recipient_webgiverinc.html',
        f'{pwd}{sep}templates{sep}giver_nowebgiver.html',
        f'{pwd}{sep}templates{sep}recipient_nowebgiver.html',
        f'{pwd}{sep}templates{sep}giver_stg_annual.html',
        f'{pwd}{sep}templates{sep}recipient_stg_annual.html',
        f'{pwd}{sep}templates{sep}giver_stg_monthly.html',
        f'{pwd}{sep}templates{sep}recipient_stg_monthly.html',
    ]

    subjects = [
        'Thank You for Your Aquarium Gift Membership Purchase!',
        'You\'ve Been Given the Gift of Membership to the South Carolina Aquarium!',
        'Thank You for Your Aquarium Gift Membership Purchase!',
        'You\'ve Been Given the Gift of Membership to the South Carolina Aquarium!',
        'Thank You for Gifting a Sea Turtle Guardianship!',
        'You\'ve Been Given the Gift of a Sea Turtle Guardianship!',
        'Thank You for Gifting a Sea Turtle Guardianship!',
        'You\'ve Been Given the Gift of a Sea Turtle Guardianship!',
    ]

    csv_file = find_first_with_ext_in_dir('csv')

    queries = [
        'MEM-Gift_Primary_Web Giver Inc_Acknowledgement Letter',
        'MEM-Gift_Giver_Web Giver Inc_Acknowledgement Letter',
        'STG-Gift_Giver_Annual_Acknowledgement Email',
        'STG-Gift_Giver_Monthly_Acknowledgement Email'
    ]

    def __init__(self, data):
        """

        :param kwargs: the data dependent on a query
        """
        self.data = data
        self.templates = list()
        self.subjects = list()
        self.csv_file = RecordData.csv_file
        self.queries = RecordData.queries  # TODO don't think this is used...

    def reset_record_data(self, query_string):

        templateObj = self.get_obj_by_query(query_string)

        self.templates = [templateObj['filename']['giver'],
                          templateObj['filename']['recipient']]
        self.subjects = [templateObj['subject']['giver'],
                         templateObj['subject']['recipient']]

        self.add_path_to_templates()

    def get_obj_by_query(self, query_string):
        for obj in self.data['templatesData']:
            if obj['queryName'] == query_string:
                return obj

    def add_path_to_templates(self):
        template_temp = list()
        for file_string in self.templates:
            template_temp.append(
                f'{RecordData.pwd}{RecordData.sep}{self.data["templatesFolder"]}{RecordData.sep}{file_string}'
            )

        self.templates = template_temp


if __name__ == '__main__':
    x = RecordData()
    x.reset_record_data(
        'MEM-Gift_Primary_Web Giver Inc_Acknowledgement Letter')
    print(x.templates[1])
    print(x.csv_file)
