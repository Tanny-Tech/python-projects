import openpyxl
import os
import emoji
import pwinput
import getpass
import time 
import uuid
import datetime

from rich.console import Console
from pyfiglet import figlet_format
from getpass import getpass 
from pwinput import pwinput
from rich.progress import track
from datetime import datetime

console = Console()

file_name = "bank_records_1.xlsx"


#-----excel file creation------
"""This function creates an excel file named bank_records_1"""
def create_excel_file():
    if not os.path.exists(file_name):
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "Bank_Records"
    
        sheet.append(["Account No", "Name", "PIN", "Transaction ID", "Transaction Type", "Amount", "Previous Balance", "Current Balance", "Date-Time"])     

        wb.save(file_name)            
        wb.close()      
        console.print("[green]Excel Data Base Created successfully[/green]")
    

#------------------------------------------
#Bank_Banner
#------------------------------------------        
def Banner_bank():
    print(emoji.emojize(":bank:"))
    print(emoji.emojize(":money_bag:"))
    print(figlet_format("PY BANK"))
    console.print("[blue] Welcome to the PY BANK [/blue]")
    print("1. Create Account\n2. Login\n3. Exit")

    while True:
        choice = input("\n\nEnter your choice : ").strip()
        if choice == "1":
            CREATE_ACCOUNT()
            break
        elif choice == "2":
            login()
            break
        elif choice == "3":
            break
        else:
            console.print("[bold red]Enter a correct value (1, 2, or 3)[/bold red]")
            Banner_bank()


#-----------------------------CREATE_ACCOUNT-------------------------------
def CREATE_ACCOUNT():
    console.print("[bold yellow] CREATE ACCOUNT [/bold yellow]")
    name = input("Enter your name: ").strip()
    
    while True:
        num = input("Enter your ACCOUNT NUMBER: ").strip()
        if not num.isdigit():
            console.print("[bold red] ENTER CORRECT ACCOUNT NUMBER [/bold red]")
            continue
        break

    while True:
        pin = pwinput("Create a 4 Digit pin  : ", mask="*").strip()
        if not pin.isdigit():
            console.print("[bold red] Enter DIGIT VALUE [/bold red]")
            continue
        elif not len(pin) == 4:
            console.print("[bold red]    4 DIGIT PIN ONLY    [/bold red]")
            continue
        else:
            for i in track(range(100), description="[cyan]Processing data..."):
                time.sleep(0.02)    
            break

    try:
        open_balance = float(input("Enter an opening balance: "))
        if open_balance < 0:
            console.print("[bold yellow] OPENING BALANCE CANNOT BE NEGATIVE [/bold yellow]")
            return
    except ValueError:
        console.print("[bold red] ENTER CORRECT VALUE [/bold red]")
        return

    transaction_id = uuid.uuid4().hex[:10]
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Preparing operational values for initial account row creation
    transaction_type = "Initial Deposit"
    amount = open_balance
    previous_balance = 0.0
    current_balance = open_balance

    
    wb = openpyxl.load_workbook(file_name)
    sheet = wb.active
    
    sheet.append([num, name, pin, transaction_id, transaction_type, amount, previous_balance, current_balance, current_time])
    
    wb.save(file_name)  
    wb.close()

    for i in track(range(100), description="[cyan]Processing Your Bank Account..."):
        time.sleep(0.02)    
        
    console.print("[green]Account saved successfully in database![/green]")

#----------------------------login------------------------------
def login():
    console.print("[green]login[/green]")
    acc_num=input("Enter your Account number:").strip()
    pin=input("Enter your pin:").strip()
    
    workbook = openpyxl.load_workbook(file_name)
    sheet=workbook.active
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0]==acc_num and row[2]==pin:
            user_name=row[1]
            break

        workbook.close()
        console.print("[green]Verifing...[/ngreen]")
        for i in track(range(100), description="[cyan]Processing Your Bank Account..."):
                time.sleep(0.02)  
        console.print("Login successfull")
        console.print(f"[green]Welcome back {user_name}[/green]")
        print()
    else:
        console.print("[red]Invalid accont number[/red]")
#--------------------banking_menu--------------------------
def banking_menu(account_no):
    #
    while True:

        console.print(f"\n[bold cyan]BANKING OPTIONS for {account_no} [/bold cyan]")

        print("1. Check Balance")
        print("2. Deposit Money")
        print("3. Withdraw Money")
        print("4. Transaction History")
        print("5. Logout")

        choice = input("\nEnter Choice: ")

        if choice == "1":
            pass
        elif choice == "2":
            pass
        elif choice== "3":
            pass
        elif choice == "4":
           pass
        elif choice == "5":
            console.print("\n[yellow]Logged out successfully.[/yellow]")
            break
        else:
            console.print("[red]Invalid Choice[/red]")
                
            





























# ---------------------------MAINCODE----------------------------
create_excel_file()
Banner_bank()