import os
import csv
import pypff

def processMessage(message, folder_name, password):
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
    next_word = ''
    try:
        my_string = my_string.decode()
        string_list = my_string.split()
        for s in range(len(string_list)):
            if "Password" in string_list[s] or "password" in string_list[s] or "PWD" in string_list[s] or "pwd" in string_list[s]:
                next_word = string_list[s+1]
    except:
        next_word = ''
    return next_word

pst = pypff.file()
pst.open("test.pst")
base = pst.get_root_folder()
header = ['folder_name', 'sender', 'password']
create_csv()
folderTraverse(base)
pst.close()