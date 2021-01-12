import os
import csv
import pypff

def processMessage(message, folder_name, password):
    # return {
    #     "folder_name": folder_name,
    #     "subject": message.subject,
    #     "sender": message.sender_name,
    #     "header": message.transport_headers,
    #     "password": password,
    # }
    return {
        "folder_name": folder_name,
        "sender": message.sender_name,
        "password": password,
    }

def folderTraverse(base):
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
    with open('pst.csv', 'a+', encoding='utf-8') as f:
        csv_writer = csv.DictWriter(f, delimiter='|', fieldnames=header)
        csv_writer.writerows(message_list)
        f.close()

def create_csv():
    with open('pst.csv', 'w', encoding='utf-8') as f:
        csv_writer = csv.DictWriter(f, delimiter='|', fieldnames=header)
        csv_writer.writeheader()
        f.close()

def get_password(my_string):
    try:
        # split_string = my_string.split()
        # if 'http' in split_string:
        #     next_word = 'found'
        # else:
        #     next_word = 'not found'
        next_word = my_string.split('http', maxsplit=1)[-1].split(maxsplit=1)[0]
    except:
        next_word = ''
    return next_word

pst = pypff.file()
pst.open("test.pst")
base = pst.get_root_folder()
#header = ['folder_name', 'subject', 'sender', 'header', 'password']
header = ['folder_name', 'sender', 'password']
create_csv()
folderTraverse(base)
pst.close()