import os
import re
import sys
import csv
import glob
import tqdm
import time
import pypff
import argparse
import datefinder
from datetime import date

__author__ = 'Allan Abendanio'
__date__ = '20210111'
__version__ = 0.01
__description__ = 'This scripts handles processing and output of PST Email Containers'

def processMessage(message, folder_name, key_word, date_found, cc_found, possible_id):
    """
    The processMessage function processes multi-field messages to simplify collection of information
    :param message: pypff.Message object
    :return: A dictionary with message fields (values) and their data (keys)
    """
    return {
        "folder_name": folder_name,
        "sender": message.sender_name,
        "key_word": key_word,
        "date_found": date_found,
        "cc_found": cc_found,
        "possible_id": possible_id,
    }

def folderTraverse(base, file, desc):
    """
    The folderTraverse function walks through the base of the folder and scans for sub-folders and messages
    :param base: Base folder to scan for new items within the folder.
    :return: None
    """
    for folder in tqdm.tqdm(base.sub_folders, desc=desc, unit= " " + str(file), ncols= 100):
        time.sleep(.01)
        if folder.number_of_sub_folders:
            folderTraverse(folder, file, desc="Parsing Sub Folders")
        message_list = []
        for message in folder.sub_messages:
            password = get_keyword_nextword(message.plain_text_body)
            key_word2 = [x for x in key_word if x]
            key_word2 = ", ".join(str(e) for e in key_word2)
            date_found2 = [x for x in date_found if x]
            date_found2 = ", ".join(str(e) for e in date_found2)
            cc_found2 = [x for x in cc_found if x]
            cc_found2 = ", ".join(str(e) for e in cc_found2)
            possible_id2 = [x for x in possible_id if x]
            possible_id2 = ", ".join(str(e) for e in possible_id2)
            #print(str(len(key_word2)), str(len(date_found2)), str(len(cc_found2)), str(len(possible_id2)))
            message_dict = processMessage(message, folder.name, key_word2, date_found2, cc_found2, possible_id2)
            message_list.append(message_dict)
        folderReport(message_list)

def folderReport(message_list):
    """
    The folderReport function generates a report per PST folder
    :param message_list: A list of messages discovered during scans
    :folder_name: The name of an Outlook folder within a PST
    :return: None
    """
    with open(filename, 'a+', encoding='utf-8') as f:
        csv_writer = csv.DictWriter(f, delimiter='|', fieldnames=header)
        csv_writer.writerows(message_list)
        f.close()

def get_keyword_nextword(my_string):
    global key_word, date_found, cc_found, possible_id
    possible_id = []
    key_word = []
    date_found = []
    cc_found = []
    try:
        my_string = my_string.decode()
        string_list = my_string.split()
        for s in range(len(string_list)):
            for k in range(len(keywords)):
                if keywords[k].lower() in string_list[s].lower():
                    if 'address' in string_list[s].lower()  or 'addrs' in string_list[s].lower()  or 'addr' in string_list[s].lower():
                        key_word.append(string_list[s] + " " + string_list[s+1] + " " + string_list[s+2] + " " + string_list[s+3] + " " + string_list[s+4] + " " + string_list[s+5] + " " + string_list[s+6] + " " + string_list[s+7])
                    elif string_list[s+1].lower() != 'is':
                        key_word.append(string_list[s] + " " + string_list[s+1])
                    else:
                        key_word.append(string_list[s] + " " + string_list[s+2])
            df = list(datefinder.find_dates(string_list[s]))
            try:
                date_found.append(df[0].strftime('%m/%d/%Y'))
            except:
                pass
            cc = re.findall(r"(^|\s+)(\d{4}[ -]\d{4}[ -]\d{4}[ -]\d{4})(?:\s+|$)", string_list[s])
            cc_found.append(cc)
            id = re.findall(r'(?:\d+[a-zA-Z]+|[a-zA-Z]+\d+)', string_list[s])
            possible_id.append(id)
    except Exception as e:
        pass
    return key_word, date_found, cc_found, possible_id

def get_keywords():
    if not os.path.exists('keywords.txt'):
	    sys.exit("The keywords.txt not found, place it under same directory of application, please try again.")
    global keywords
    keywords = []
    with open('keywords.txt', 'r', encoding='utf-8') as csv_file:
        csv_reader = csv.DictReader(csv_file, delimiter=',')
        for row in tqdm.tqdm(csv_reader, desc="Extracting", unit=" keywords.txt"):
            time.sleep(.01)
            keyword = row['keyword']
            keywords.append(keyword)
    csv_file.close()

def create_csv(file):
    global filename
    filename = "report/" + str(file)[:-4] + "_" + date.today().strftime('%Y.%m.%d') + ".csv"
    file_exist = False
    file_int = 0
    if not os.path.exists(filename):
        os.makedirs(os.path.dirname(filename), 0o777, True)
        data = open(filename, "w", encoding='utf-8')
        file_exist = True
    while (file_exist == False):
        file_int += 1
        filename = "report/" + str(file)[:-4] + "_" + date.today().strftime('%Y.%m.%d') + "_" + str(file_int) + ".csv"
        if not os.path.exists(filename):
            os.makedirs(os.path.dirname(filename), 0o777, True)
            data = open(filename, "w", encoding='utf-8')
            file_exist = True
    with open(filename, 'w', encoding='utf-8') as f:
        csv_writer = csv.DictWriter(f, delimiter='|', fieldnames=header)
        csv_writer.writeheader()
        f.close()

if __name__ == "__main__":
    get_keywords()
    pst = pypff.file()
    if glob.glob('*.pst'):
        for file in glob.glob("*.pst"):
            pst.open(file)
            base = pst.get_root_folder()
            header = ['folder_name', 'sender', 'key_word', 'date_found', 'cc_found', 'possible_id']
            create_csv(file)
            folderTraverse(base, file, desc="Parsing Folders")
            pst.close()
            for folder in tqdm.trange(100, desc="Saving", unit= " " + str(filename), ncols= 100):
                time.sleep(.01)
                pass
            print(f"Total: Folder: 0, Message: 0, Keyword matched: {len(key_word)}, Date found: {len(date_found)}, Credit Card found: {len(cc_found)}, Posible IDs: {len(possible_id)}.")                
    else:
        sys.exit("The pst file/s not found, place it under same directory of application, please try again.")
    print("Done")
    
