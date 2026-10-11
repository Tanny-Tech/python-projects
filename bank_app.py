import openpyxl
import os
import emoji
import pwinput
import getpass





from rich.console import Console
from pyfiglet import figlet_format
from getpass import getpass 
from pwinput import pwinput



console=Console()






file_name="bank_records_1.xlsx"


#-----excel file creation------
"""This function creates an excel file named bank_records_2"""
def create_excel_file():
    if not os.path.exists(file_name):
            
        wb=openpyxl.Workbook()
        sheet=wb.active
        sheet.title = "Bank_Records"
        headers=["Account No", "Name", "PIN", "Transaction ID", "Transaction Type", "Amount", "Previous Balance", "Current Balance", "Date-Time"]
        sheet.append(headers)     


        wb.save(file_name)            
        wb.close()      
        console.print("[green]Excel data Base Created successfully[/green]")
    


#------------------------------------------
#Bank_Banner
#------------------------------------------        

def Banner_bank():
    print(emoji.emojize(":bank:"))
    print(emoji.emojize(":money_bag:"))
    print(figlet_format("PY BANK"))
    console.print("[blue] Welcome to the PY BANK [/blue]")
    print("1. Create Account\n2. Login\n3. Exit")

    choice=(input("\n\nEnter your choice :"))

    while True :
        if choice=="1":
            CREATE_ACCOUNT()
            break
            
        elif choice=="2":
            print("wait for some more time")
            break
            
        elif choice=="3":
            break
        else:
            print("Enter correct value")
            Banner_bank()

#-----------------------------CREATE_ACCOUNT-------------------------------
def CREATE_ACCOUNT():
    console.print("[bold yellow] CREATE ACCOUNT [/bold yellow]")
    name=input("Enter your name:").strip()
    num=input("Enter your ACCOUNT NUMBER:")
    while True:

        if not num.isdigit():
            console.print("[bold red] ENTER CORRECT ACCOUNT NUMBER [/bold red]")
            CREATE_ACCOUNT()
            break


        elif num == "exists":
                pass

        
        pin=pwinput("Create a 4 Digit pin  : ",mask="*")
        if not pin.isdigit():
                console.print("[bold red] Enter DIGIT VALUE [/bold red]")
                break
        elif not len(pin) == 4 :
            console.print("[bold red]    4 DIGIT PIN ONLY    [/bold red]")
            break
        else:
            break


            
# ---------------------------MAINCODE----------------------------
create_excel_file()
Banner_bank()


                               
# Banner_bank()
# if __name__ == "__main__":

#     create_excel_file()


