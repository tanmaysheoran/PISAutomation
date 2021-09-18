from webaccess import *
import time

print("""  _____  _____   _____                   _                            _    _                __      __ __ 
 |  __ \|_   _| / ____|     /\          | |                          | |  (_)               \ \    / //_ |
 | |__) | | |  | (___      /  \   _   _ | |_  ___   _ __ ___    __ _ | |_  _   ___   _ __    \ \  / /  | |
 |  ___/  | |   \___ \    / /\ \ | | | || __|/ _ \ | '_ ` _ \  / _` || __|| | / _ \ | '_ \    \ \/ /   | |
 | |     _| |_  ____) |  / ____ \| |_| || |_| (_) || | | | | || (_| || |_ | || (_) || | | |    \  /_   | |
 |_|    |_____||_____/  /_/    \_\\__,_| \__|\___/ |_| |_| |_| \__,_| \__||_| \___/ |_| |_|     \/(_)  |_|
                                                                                                          
                                                                                                          """)

print()

print("Hi There, this is the first version of PIS Automation. \n\n\n**Read Me** \n--> Chromedriver is a must have to run this application and is included with the package.\n--> The locations for both chromedriver and excel file are fixed at this point but will be editable in coming versions.\n--> The script updates the daily time card, all you have to do is update the excel sheet daily.\n--> Save it and run this script.\n\n\n Thank you and enjoy!!! \n - Tanmay Sheoran")

print(50 * ("-"))
print("Let's Begin")
print("\n\n")
while True:
    print("At any time if you want to quit, enter Q")
    print()
    excel_update = input("Is your excel updated? Y/N: ")
    if excel_update == 'y' or excel_update == "Y":
        print("Great!")
        print()
        start_script = input("Start the script and chill? Y/N: ")
        if start_script == 'y' or start_script == "Y":
            selenium_script_add_values()
            print("Everything updated now, Bbye")
            time.sleep(3)
            break
        elif start_script == 'q' or start_script == "Q":
            print("Bye")
            time.sleep(3)
            break
        else:
            print("Alright, take your time :)")
    elif excel_update == 'q' or excel_update == "Q" :
        print("Bye")
        time.sleep(3)
        break
    else:
        print("What are you doing here?...")


