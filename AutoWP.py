#!/usr/bin/python3
# Description : This is a Python script that automatically sends Whatsapp messages to users contained in an Excel file.
# Author : yb0zkurt
# Date : 28.08.2021

import pywhatkit as py
import xlrd2
import argparse
import sys
from colorama import init, Fore

init()

# Constants
user_phone_list = []
user_phone_dict = {}
excelFile = ""
shour = 0
sminute = 0


# Parsing/Checking arguments
def getargv():
    try:
        global excelFile, shour, sminute
        parser = argparse.ArgumentParser(add_help=False)
        parser.add_argument("--help", "-h", action='store_true', help=argparse.SUPPRESS)
        parser.add_argument("--file", "-F", help="Set file name")
        parser.add_argument("--startHour", "-H", help="Set start hour value")
        parser.add_argument("--startMinute", "-M", help="Set start minute value")
        args = parser.parse_args()

        if args.help:
            print("""Usage: python3 AutoWP.py <excel_file> <start_hour> <start_minute>
            -h, --help : Display usage
            -F, --file : Excel file name
            -H, --startHour : Start hour
            -M, --startMinute : Start minute
            Example : python3 AutoWP.py -F users.xlsx -H 23 -M 45""")
            sys.exit(0)

        excelFile = args.file
        shour = int(args.startHour)
        sminute = int(args.startMinute)
        print(Fore.GREEN + "[Successful] AutoWP is starting...")

    except:
        print(Fore.RED + "[Error] Error! Arguments did not parse!")
        sys.exit(0)


# Reading usernames, phone numbers, and messages from excel file (AutoWP_DB.xlsx)
def getUserInfo(file):
    try:
        print(Fore.CYAN + "[Info] Parsing excel file...")
        wb = xlrd2.open_workbook(file)
        sheet = wb.sheet_by_index(0)
        for x in range(sheet.nrows):
            user_phone_dict = {"user": sheet.cell_value(x, 0), "phone": sheet.cell_value(x, 1),
                               "msg": sheet.cell_value(x, 2)}
            user_phone_list.append(user_phone_dict)
        print(Fore.GREEN + "[Successful] " + str(len(user_phone_list)) + " user's info found!")
        print(Fore.GREEN + "[Successful] Excel file parsed successfully!")
    except:
        print(Fore.RED + "[Error] Error! Excel file not found!")
        sys.exit(0)


# Sending messages
def sendingMsg(a, b):
    try:
        print(Fore.CYAN + "[Info] Sending message...")
        counter = 0
        sh = int(a)
        sm = int(b)
        for each_dict in user_phone_list:
            final_number = "+" + str((each_dict['phone']))[0:-2]
            print("\nSending message to --> " + each_dict['user'] + " - " + final_number + " - Message : " + each_dict[
                'msg'])
            py.sendwhatmsg(final_number, each_dict['msg'], sh, sm, 10, True, 5)
            if counter < len(user_phone_list):
                if sm == 59:
                    sm = 0
                else:
                    sm += 1
        print(Fore.GREEN + "[Successful] Messages sent successfully!")
    except:
        print(Fore.RED + "[Error] Error! Messages didn't send!")
        sys.exit(0)


if __name__ == '__main__':
    getargv()
    getUserInfo(excelFile)
    sendingMsg(shour, sminute)