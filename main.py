import shutil 
import time 
import pyautogui
import os
import re
from functions.generate_report_functions import * 
from classes.database_access import DB_Connect

my_db = DB_Connect("user", "password", "database")

mergeMastersAddress = "\\\PC\\path_to_file"

config_contents_list = convert_txt_to_list()
config_contents_list_formatted = format_config_list(config_contents_list)
mergeMasters, records = seperate_mergemasters_and_functions(config_contents_list_formatted)

counter = 0
print(config_contents_list, config_contents_list_formatted, mergeMasters, records)
for mergeMaster in mergeMasters:
    current_records = []
    for record in records[counter]:
        current_records.append(record)
    for current_record in current_records:
        active_file = open_mergemaster(mergeMaster[counter])
        create_report(current_record, ".odt")
        save_report_as_pdf(current_record)

        number_of_items = find_number_of_items(my_db, record)
        client = find_client(my_db, record)
        has_items = open_receipt(number_of_items, client, current_record) #change client2 to 1 when complete
        if has_items:
            create_report(current_record, ".pdf", is_pdf=True, is_receipt=True)
        date = format_current_date()
        print("Date output to DATEf will be", date)
        
    print("----------")
    counter += 1
