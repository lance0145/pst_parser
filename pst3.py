import os
import re
import csv
import pypff
import datefinder

__author__ = 'Allan Abendanio'
__date__ = '20210111'
__version__ = 0.01
__description__ = 'This scripts handles processing and output of PST Email Containers'

def processMessage(message, folder_name, password):
    """
    The processMessage function processes multi-field messages to simplify collection of information
    :param message: pypff.Message object
    :return: A dictionary with message fields (values) and their data (keys)
    """
    return {
        "folder_name": folder_name,
        "sender": message.sender_name,
        "password": password,
    }

def folderTraverse(base):
    """
    The folderTraverse function walks through the base of the folder and scans for sub-folders and messages
    :param base: Base folder to scan for new items within the folder.
    :return: None
    """
    for folder in base.sub_folders:
        if folder.number_of_sub_folders:
            folderTraverse(folder)
        message_list = []
        for message in folder.sub_messages:
            password = get_password(message.plain_text_body)
            message_dict = processMessage(message, folder.name, password)
            message_list.append(message_dict)
        folderReport(message_list)

def folderReport(message_list):
    """
    The folderReport function generates a report per PST folder
    :param message_list: A list of messages discovered during scans
    :folder_name: The name of an Outlook folder within a PST
    :return: None
    """
    with open('pst.csv', 'a+', encoding='utf-8') as f:
        csv_writer = csv.DictWriter(f, delimiter='|', fieldnames=header)
        csv_writer.writerows(message_list)
        f.close()

def create_csv():
    with open('pst.csv', 'w', encoding='utf-8') as f:
        csv_writer = csv.DictWriter(f, delimiter='|', fieldnames=header)
        csv_writer.writeheader()
        f.close()

def get_keyword_nextword(my_string):
    key_word = []
    next_word = []
    date_found = []
    cc_found = []
    try:
        my_string = my_string.decode()
        string_list = my_string.split()
        for s in range(len(string_list)):
            for k in range(len(keywords)):
                if keywords[k] in string_list[s]:
                    key_word.append(string_list[s])
                    next_word.append(string_list[s+1])
            date_found.append(datefinder.find_dates(string_list[s]))
            cc_found.append(re.search(r"(^|\s+)(\d{4}[ -]\d{4}[ -]\d{4}[ -]\d{4})(?:\s+|$)", string_list[s]))
    except:
        pass
    return keyword, next_word, date_found, cc_found

def get_keyword():
    keywords = []
    with open('input.txt', 'r', encoding='utf-8') as csv_file:
        csv_reader = csv.DictReader(csv_file, delimiter=',')
        for row in csv_reader:
            keyword = row['keyword']
            print(f'Extracting {keyword}')
            keywords.append(keyword)
    csv_file.close()

if __name__ == "__main__":
    get_keyword()
    pst = pypff.file()
    pst.open("test.pst")
    base = pst.get_root_folder()
    header = ['folder_name', 'sender', 'password']
    create_csv()
    folderTraverse(base)
    pst.close()