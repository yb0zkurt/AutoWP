# AutoWP
This is a Python script that automatically sends Whatsapp messages to users contained in an Excel file.

Usage: python3 AutoWP.py -F <excel_file> -H <start_hour> - M <start_minute>
            
            -h, --help : Display usage

            -F, --file : Excel file name

            -H, --startHour : Start hour

            -M, --startMinute : Start minute

            Example : python3 AutoWP.py -F users.xlsx -H 23 -M 45


*** Install requirements.txt

*** Add user's names, phone numbers, and messages in AutoWP_DB.xlsx file.

    Note: Add phone number without "+" character in AutoWP_DB.xlsx file. 