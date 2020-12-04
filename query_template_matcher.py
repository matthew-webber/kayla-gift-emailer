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
    templates = [
        f'{pwd}{sep}templates{sep}v1_giver_template.html',
        f'{pwd}{sep}templates{sep}v1_recipient_template.html',
        f'{pwd}{sep}templates{sep}v2_giver_template.html',
        f'{pwd}{sep}templates{sep}v2_recipient_template.html',
        f'{pwd}{sep}templates{sep}sea_turtle_giver.html',
        f'{pwd}{sep}templates{sep}sea_turtle_recipient.html',
    ]

    subjects = [
        'Thank You for Your Aquarium Gift Membership Purchase!',
        'You\'ve Been Given the Gift of Membership to the South Carolina Aquarium!',
        'Thank You for Your Aquarium Gift Membership Purchase!',
        'You\'ve Been Given the Gift of Membership to the South Carolina Aquarium!',
        'Thank You for Gifting a Sea Turtle Guardianship!',
        'You\'ve Been Given the Gift of a Sea Turtle Guardianship!',
    ]
    
    csv_file = find_first_with_ext_in_dir('csv')
    
    queries = [
        'MEM-Gift_Primary_Web Giver Inc_Acknowledgement Letter',  # Mem v1
        'MEM-Gift_Giver_Web Giver Inc_Acknowledgement Letter',  # mem v2
        'STG-Gift_Giver_Annual_Acknowledgement Email',
    ]

    def __init__(self, **kwargs):
        """

        :param kwargs: the data dependent on a query
        """
        self.templates = None
        self.subjects = None
        self.csv_file = RecordData.csv_file
        self.queries = RecordData.queries
        self.query_matcher = self.create_query_matcher()

    def reset_record_data(self, query_string):
        self.templates = self.query_matcher[query_string]['templates']
        self.subjects = self.query_matcher[query_string]['subjects']

    def create_query_matcher(self):
        return {
            self.queries[0]: {
                'templates': [
                    RecordData.templates[0],
                    RecordData.templates[1],
                ],
                'subjects': [
                    RecordData.subjects[0],
                    RecordData.subjects[1],
                ]
            },
            self.queries[1]: {
                'templates': [
                    RecordData.templates[2],
                    RecordData.templates[3],
                ],
                'subjects': [
                    RecordData.subjects[2],
                    RecordData.subjects[3],
                ]
            },
            self.queries[2]: {
                'templates': [
                    RecordData.templates[4],
                    RecordData.templates[5],
                ],
                'subjects': [
                    RecordData.subjects[4],
                    RecordData.subjects[5],
                ]
            }
        }


if __name__ == '__main__':
    x = RecordData()
    x.reset_record_data('MEM-Gift_Primary_Web Giver Inc_Acknowledgement Letter')
    print(x.templates[1])
    print(x.csv_file)