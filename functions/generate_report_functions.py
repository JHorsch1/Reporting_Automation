import shutil 
import time
from numpy import number 
import pyautogui
import re
import os
from datetime import datetime
from classes.database_access import DB_Connect

def record_input():
    """Function to take one or many record numbers from the user and return them as a list"""
    recordInput = pyautogui.prompt("Enter the record number of report. If multiple, seperate each record with a comma (no space)")
    recordList = recordInput.split(",")
    return recordList

def format_current_date():
    """Function to take current date, and format it into a string that matches the format of the database"""
    current_date = datetime.now()
    current_date_formatted = current_date.strftime("%y-%d-%m")
    current_date_formatted_string = str(current_date_formatted)
    return current_date_formatted_string

def format_current_time():
    """Function to take current time, and format it into a string that matches the format of the database"""
    current_date = datetime.now()
    current_time_formatted = current_date.strftime("%H:%M:%S")
    current_time_formatted_string = str(current_time_formatted)
    return current_time_formatted_string

def output_datef_to_database(my_db, record):
    """Function to output current date to datef in the database"""
    current_date_formatted_string = format_current_date()
    date_query = my_db.executeQuery(F"UPDATE main SET DATEf = '{current_date_formatted_string}' WHERE record = {record}")
    my_db.conn.commit()

def output_current_time_and_date_to_database(my_db, record_number):
    """Function to take current time and date along with the current record, and output to the database"""
    current_date_formatted_string = format_current_date()
    current_time_formatted_string = format_current_time()
    date_time_query = my_db.executeQuery(f"UPDATE main SET DATEf = '{current_date_formatted_string}', TIMEf = '{current_time_formatted_string}' WHERE record = {record_number}")
    my_db.conn.commit()

def convert_txt_to_list():
    """Function to read the txt doc containing form names and record numbers, and convert to a string"""     
    config_contents_list = []
    with open("text_files/config.txt", "r") as config:
        config_contents = config.readlines()
    for line in config_contents:
        config_contents_list.append(line)
    return config_contents_list

def format_config_list(config_contents_list):
    """Function to re-format the config list and remove newlines and empty elements"""
    config_contents_list_formatted = []
    for index in config_contents_list:
        config_contents_list_formatted.append(index.strip())
    #config_contents_list_formatted.pop()
    return config_contents_list_formatted

def seperate_mergemasters_and_functions(config_contents_list_formatted):
    """Function to return mergemaster names and associated records in seperate lists"""
    mergemasters = []
    records = []
    for index in config_contents_list_formatted:
        mergemaster_regex = re.findall(r'^[^(:]*[^(:\s]', index)
        records_regex = re.findall(r'\d{1,4}', index)
        mergemasters.append(mergemaster_regex)
        records.append(records_regex)
    return mergemasters, records

def open_mergemaster(mergeMaster):
    """Function to open the merge master file using the passed filename"""
    mergeMastersAddress = "\\\PC_NAME\\Path_to_file\\"
    active_file = os.startfile(mergeMastersAddress + mergeMaster + ".odt")
    return active_file

def check_for_authentication():
    """Function to sign in to the MySQL account if needed"""
    try:
        authenticationRequired = pyautogui.locateOnScreen("images/authenticationRequired.png", confidence = 0.9)
        authenticationRequiredTest = authenticationRequired[0] + 1
    except:
        print("Authentication not required")
    else:
        authenticationRequiredPassword = pyautogui.locateOnScreen("images/authenticationRequiredPassword.png", confidence = 0.9)
        authenticationRequiredPasswordX = authenticationRequiredPassword[0] + 100
        authenticationRequiredPasswordY = authenticationRequiredPassword[1]
        pyautogui.click(authenticationRequiredPasswordX, authenticationRequiredPasswordY)
        pyautogui.typewrite("password")
        pyautogui.press("enter")
        time.sleep(5)

def create_report(record, format, is_pdf=False, is_receipt=False):
    """Function to use macros to create report based on what mergemaster is currently open/visible on screen"""
    report_save_address = "\\\PC_NAME\\Path_to_file\\"

    print(record)
    pyautogui.sleep(7)
    printLocation = pyautogui.locateOnScreen("images/writerPrint.png", confidence = 0.9)
    pyautogui.click(printLocation)
    printLocation = pyautogui.locateOnScreen("images/writerPrint.png", confidence = 0.9)
    pyautogui.click(printLocation)
    time.sleep(3)
    printYesLocation = pyautogui.locateOnScreen("images/printYes.png", confidence = 0.9)
    pyautogui.click(printYesLocation)
    time.sleep(5)

    check_for_authentication()

    time.sleep(5)
   
    printFromIntLocation = pyautogui.locateOnScreen("images/printFromInt.png", confidence = 0.9)
    printFromIntLocationX = printFromIntLocation[0] + 80
    printFromIntLocationY = printFromIntLocation[1] + 15
    pyautogui.click(printFromIntLocationX, printFromIntLocationY)
    pyautogui.press("backspace")
    pyautogui.typewrite(record)
    printToIntLocation = pyautogui.locateOnScreen("images/printToInt.png", confidence = 0.9)
    pyautogui.click(printToIntLocation)
    pyautogui.press("backspace")
    pyautogui.typewrite(record)

    printFileSelect = pyautogui.locateOnScreen("images/printFileSelect.png", confidence = 0.9)
    pyautogui.click(printFileSelect)

    if is_receipt:
        noOpenQuery()
    if is_receipt:
        try:
            invoiceSelected = pyautogui.locateOnScreen("images/invoiceQuerySelected.png", confidence = 0.9)
            invoiceSelectedTest = invoiceSelected[0] + 1
        except:
            tableSelectorDown = pyautogui.locateOnScreen("images/tableSelectorDown.png", confidence = 0.9)
            pyautogui.click(tableSelectorDown)
            pyautogui.click(tableSelectorDown)
            invoiceQueryUnselected = pyautogui.locateOnScreen("images/invoiceQueryUnselected.png", confidence = 0.9)
            pyautogui.click(invoiceQueryUnselected)
            pyautogui.PAUSE = 3
        
        select_group_of_records(record)

    pyautogui.press("enter")
    pyautogui.PAUSE = 3

    pyautogui.keyDown("ctrl")
    pyautogui.press("l")
    pyautogui.keyUp("ctrl")
    pyautogui.typewrite(report_save_address)
    pyautogui.press("enter")

    if is_pdf or is_receipt:
        saveAsType = pyautogui.locateOnScreen("images/saveAsType.png", confidence = 0.9)
        pyautogui.click(saveAsType)
        pyautogui.press("pagedown")
        pyautogui.press("up")
        pyautogui.press("enter")

    pyautogui.PAUSE = 3
    fileNameBox = pyautogui.locateOnScreen("images/fileNameBox.png")
    pyautogui.click(fileNameBox)
    if is_pdf:
        pyautogui.typewrite(record + format)
    else:
        pyautogui.typewrite(record + format)
    pyautogui.press("enter")
    time.sleep(2)

    #Adding to log 
    if is_pdf:
        add_to_log(record, is_report_pdf=True)
    elif is_receipt:
        add_to_log(record, is_receipt=True)
    else:
        add_to_log(record, is_report_odt=True)

    #closing the file
    pyautogui.keyDown("alt")
    pyautogui.press("f4")
    pyautogui.keyUp("alt")

def noOpenQuery():
    """Runs if the tables option is expanded instead of queries"""
    try:
        invoiceExpand = pyautogui.locateOnScreen("images/find_query/invoiceExpand.png", confidence = 0.9)
        pyautogui.click(invoiceExpand)
    except:
        pass

def noQuerySelected():
    """Runs if no query is actively selected while generating invoice"""
    try:
        noQuerySelected = pyautogui.locateOnScreen("images/noSelectedQuery.png", confidence = 0.9)
        pyautogui.click(noQuerySelected)
    except:
        pass
    else:
        queryFolder = pyautogui.locateOnScreen("images/queryFolder.png", confidence = 0.9)
        pyautogui.click(queryFolder)
        counter = 0
        time.sleep(1)
        while counter < 5:
            time.sleep(1)
            pyautogui.press("down")
            counter = counter + 1

def select_group_of_records(record):
    """Searches and selctes multiple rows in the invoice query if multiple items are associated with the record"""
    my_db = DB_Connect("user", "password", "database")
    number_of_items = find_number_of_items(my_db, record)
    number_of_items = int(number_of_items)
    time.sleep(2)
    #noQuerySelected()#checks if invoice query is selected 
    searchIcon = pyautogui.locateOnScreen("images/searchIcon.png", confidence = 0.9)
    pyautogui.click(searchIcon)
    time.sleep(4)
    pyautogui.typewrite(record)
    pyautogui.press("enter")
    # searchBox = pyautogui.locateOnScreen("images/searchBoxSelect.png", confidence = 0.9)
    # pyautogui.click(searchBox)
    time.sleep(4)
    closeBox = pyautogui.locateOnScreen("images/closeBox.png", confidence = 0.9)
    pyautogui.click(closeBox)
    time.sleep(2)
    selectedRecord = pyautogui.locateOnScreen("images/selectedRecordArrow.png", confidence = 0.9)
    pyautogui.click(selectedRecord)
    time.sleep(2)
    pyautogui.keyDown("shiftleft")
    pyautogui.keyDown("shiftright")
    time.sleep(2)
    counter = 1
    while counter < number_of_items:
        time.sleep(2)
        pyautogui.press("down")
        counter = counter + 1
    pyautogui.keyUp("shiftleft")
    pyautogui.keyUp("shiftright")

def find_number_of_items(my_db, record):
    """Function to find the number of items to decide what receipt to open for a given record"""
    number_of_items = my_db.executeSelectQuery(f"SELECT count(*) FROM sold where INVOICE_ID = {record}")
    str_items = str(number_of_items[0])
    items_regex = re.findall(r'\d{1,4}', str_items)
    int_item = items_regex[0]
    return int_item

def find_client(my_db, record):
    """Function to find the client of a given record"""
    client = my_db.executeSelectQuery(f"SELECT client2 FROM main where Record = {record}")
    client_string = client[0]
    client_string = str(client_string)
    client_string = client_string[:-2]
    client_string = client_string[13:]
    return client_string

def open_receipt(number_of_items, client, current_record):
    """Function to find and open the correct receipt depending on the number of items and client"""
    mergeMastersDir = "\\\PC_NAME\\Path_to_file\\"
    number_of_items = int(number_of_items) #number_of_items is initially passed in as a string
    if number_of_items:
        if client == "EXAMPLE":
            if number_of_items == 1:
                os.startfile(mergeMastersDir + "EXAMPLE INVOICE MASTER 1 ITEM.odt")
            elif number_of_items == 2:
                os.startfile(mergeMastersDir + "EXAMPLE INVOICE MASTER 2 ITEMS.odt")
            elif number_of_items == 3:
                os.startfile(mergeMastersDir + "EXAMPLE INVOICE MASTER 3 ITEMS.odt")
            elif number_of_items == 4:
                os.startfile(mergeMastersDir + "EXAMPLE INVOICE MASTER 4 ITEMS.odt")
            elif number_of_items == 5:
                os.startfile(mergeMastersDir + "EXAMPLE INVOICE MASTER 5 ITEMS.odt")
        elif client == "EXAMPLE2":
            if number_of_items == 1:
                os.startfile(mergeMastersDir + "EXAMPLE2 INVOICE MASTER 1 ITEM.odt")
            elif number_of_items == 2:
                os.startfile(mergeMastersDir + "EXAMPLE2 INVOICE MASTER 2 ITEMS.odt")
            elif number_of_items == 3:
                os.startfile(mergeMastersDir + "EXAMPLE2 INVOICE MASTER 3 ITEMS.odt")
            elif number_of_items == 4:
                os.startfile(mergeMastersDir + "EXAMPLE2 INVOICE MASTER 4 ITEMS.odt")
            elif number_of_items == 5:
                os.startfile(mergeMastersDir + "EXAMPLE2 INVOICE MASTER 5 ITEMS.odt")
        else:
            if number_of_items == 1:
                os.startfile(mergeMastersDir + "GENERIC_EXAMPLE MASTER 1 ITEM.odt")
            elif number_of_items == 2:
                os.startfile(mergeMastersDir + "GENERIC_EXAMPLE 2 ITEMS.odt")
            elif number_of_items == 3:
                os.startfile(mergeMastersDir + "GENERIC_EXAMPLE 3 ITEMS.odt")
            elif number_of_items == 4:
                os.startfile(mergeMastersDir + "GENERIC_EXAMPLE 4 ITEMS.odt")
            elif number_of_items == 5:
                os.startfile(mergeMastersDir + "GENERIC_EXAMPLE 5 ITEMS.odt")
            elif number_of_items == 6:
                os.startfile(mergeMastersDir + "GENERIC_EXAMPLE 6 ITEMS.odt")
            elif number_of_items == 7:
                os.startfile(mergeMastersDir + "GENERIC_EXAMPLE 7 ITEMS.odt")
            elif number_of_items == 8:
                os.startfile(mergeMastersDir + "GENERIC_EXAMPLE 8 ITEMS.odt")
            elif number_of_items == 9:
                os.startfile(mergeMastersDir + "GENERIC_EXAMPLE 9 ITEMS.odt")
            elif number_of_items == 10:
                os.startfile(mergeMastersDir + "GENERIC_EXAMPLE 10 ITEMS.odt")
        return True
    else:
        print(f"No items for {current_record} - output to log")
        with open('logs/log.txt', 'a') as log:
            log.write(f"Record {current_record} - No associated items\n")
        return False

def save_report_as_pdf(current_record):
    """Function to open the last created odt report, and save it as pdf"""
    merged_reports_address = "\\\PC_NAME\\Path_to_file\\"
    counter = 0
    os.startfile(merged_reports_address + current_record + ".odt") #remove test for final release
    time.sleep(5)
    pyautogui.keyDown("ctrl")
    pyautogui.press("6")
    pyautogui.keyUp("ctrl")
    time.sleep(3)
    pyautogui.keyDown("alt")
    pyautogui.press("f4")
    pyautogui.keyUp("alt")

def add_to_log(record, is_report_odt=False, is_report_pdf=False, is_receipt=False):
    """Function to output to the log whenever a successful report/pdf is created"""
    with open('logs/log.txt', 'a') as log:
        if is_report_odt:
            log.write(f"Record {record} - odt report creation success\n")
        elif is_report_pdf:
            log.write(f"Record {record} - pdf report creation success\n")
        elif is_receipt:
            log.write(f"Record {record} - invoice creation success\n")
