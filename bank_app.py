import openpyxl
import os
import emoji
import pwinput
import getpass
import time 
import uuid
from datetime import datetime

from rich.console import Console
from pyfiglet import figlet_format
from getpass import getpass 
from pwinput import pwinput
from rich.progress import track
from rich.table import Table

console = Console()
file_name = "bank_records_1.xlsx"

#-----excel file creation------
def create_excel_file():
    """This function creates an excel file named bank_records_1"""
    if not os.path.exists(file_name):
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "Bank_Records"
    
        sheet.append(["Account No", "Name", "PIN", "Transaction ID", "Transaction Type", "Amount", "Previous Balance", "Current Balance", "Date-Time"])     

        wb.save(file_name)            
        wb.close()      
        console.print("[green]Excel Data Base Created successfully[/green]")

#-----------------------------GET BALANCE HELPER------------------------------
def get_balance(account_no):
    wb = openpyxl.load_workbook(file_name)
    sheet = wb.active
    current_balance = 0.0
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if str(row[0]) == str(account_no):
            current_balance = row[7]
    wb.close()
    return current_balance

#-----------------------------GET USER DETAILS HELPER------------------------------
def get_user_details(account_no):
    wb = openpyxl.load_workbook(file_name)
    sheet = wb.active
    name = None
    pin = None
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if str(row[0]) == str(account_no):
            name = row[1]
            pin = row[2]
            break
    wb.close()
    return name, pin

#-----------------------------CHECK BALANCE------------------------------
def check_balance(account_no):
    """Displays the current account balance."""
    balance = get_balance(account_no)
    console.print(f"\n[bold cyan]Your current balance is: ₹{balance:.2f}[/bold cyan]")

#-----------------------------DEPOSIT MONEY------------------------------
def deposit_money(account_no):
    """Allows the user to deposit money into their account."""
    console.print("\n[bold yellow]DEPOSIT MONEY[/bold yellow]")
    try:
        amount = float(input("Enter amount to deposit: ₹"))
        if amount <= 0:
            console.print("[red]Deposit amount must be greater than zero.[/red]")
            return
    except ValueError:
        console.print("[red]Invalid input! Please enter a valid number.[/red]")
        return

    previous_balance = get_balance(account_no)
    new_balance = previous_balance + amount
    name, pin = get_user_details(account_no)

    if name is None:
        console.print("[red]Account details not found.[/red]")
        return

    transaction_id = uuid.uuid4().hex[:10]
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    wb = openpyxl.load_workbook(file_name)
    sheet = wb.active
    sheet.append([account_no, name, pin, transaction_id, "Deposit", amount, previous_balance, new_balance, current_time])
    wb.save(file_name)
    wb.close()

    for i in track(range(100), description="[cyan]Verifying deposit..."):
        time.sleep(0.02)

    console.print(f"[green]Successfully deposited ₹{amount:.2f}[/green]")
    console.print(f"New Balance: [bold cyan]₹{new_balance:.2f}[/bold cyan]")

#-----------------------------WITHDRAW MONEY------------------------------
def withdraw_money(account_no):
    """Allows the user to withdraw money if sufficient balance is available."""
    console.print("\n[bold yellow]WITHDRAW MONEY[/bold yellow]")
    previous_balance = get_balance(account_no)
    
    try:
        amount = float(input("Enter amount to withdraw: ₹"))
        if amount <= 0:
            console.print("[red]Withdrawal amount must be greater than zero.[/red]")
            return
    except ValueError:
        console.print("[red]Invalid input! Please enter a valid number.[/red]")
        return

    if amount > previous_balance:
        console.print("[bold red]Insufficient balance! You cannot withdraw more than your current balance.[/bold red]")
        return

    new_balance = previous_balance - amount
    name, pin = get_user_details(account_no)
    transaction_id = uuid.uuid4().hex[:10]
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    wb = openpyxl.load_workbook(file_name)
    sheet = wb.active
    sheet.append([account_no, name, pin, transaction_id, "Withdraw", amount, previous_balance, new_balance, current_time])
    wb.save(file_name)
    wb.close()

    for i in track(range(100), description="[cyan]Processing withdrawal..."):
        time.sleep(0.02)

    console.print(f"[green]Successfully withdrew ₹{amount:.2f}[/green]")
    console.print(f"Remaining Balance: [bold cyan]₹{new_balance:.2f}[/bold cyan]")

#-----------------------------TRANSACTION HISTORY------------------------------
def transaction_history(account_no):
    """Displays recent transactions."""
    wb = openpyxl.load_workbook(file_name)
    sheet = wb.active
    table = Table(title="Transaction History")
    table.add_column("Transaction ID")
    table.add_column("Type")
    table.add_column("Amount")
    table.add_column("Previous")
    table.add_column("Balance")
    table.add_column("Date-Time")

    found = False
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if str(row[0]) == str(account_no):
            found = True
            table.add_row(str(row[3]), str(row[4]), f"₹{row[5]:.2f}", f"₹{row[6]:.2f}", f"₹{row[7]:.2f}", str(row[8]))
    
    wb.close()
    
    if found:
        console.print(table)
    else:
        console.print("[red]No transactions found.[/red]")

#-----------------------------BANKING MENU------------------------------
def banking_menu(account_no):
    while True:
        console.print(f"\n[bold cyan]BANKING OPTIONS for {account_no} [/bold cyan]")
        print("1. Check Balance")
        print("2. Deposit Money")
        print("3. Withdraw Money")
        print("4. Transaction History")
        print("5. Logout")

        choice = input("\nEnter Choice: ").strip()

        if choice == "1":
            check_balance(account_no)
        elif choice == "2":
            deposit_money(account_no)
        elif choice == "3":
            withdraw_money(account_no)
        elif choice == "4":
            transaction_history(account_no)
        elif choice == "5":  
            console.print("\n[yellow]Logged out successfully.[/yellow]")
            break
        else:
            console.print("[red]Invalid choice[/red]")

#-----------------------------CREATE_ACCOUNT------------------------------
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
        pin = pwinput("Create a 4 Digit pin : ", mask="*").strip()
        if not pin.isdigit():
            console.print("[bold red] Enter DIGIT VALUE [/bold red]")
            continue
        elif len(pin) != 4:
            console.print("[bold red]    4 DIGIT PIN ONLY    [/bold red]")
            continue
        else:
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

#----------------------------LOGIN------------------------------
def login():
    console.print("[green]---- LOGIN ----[/green]\n")
    acc_num = input("Enter your Account number: ").strip()
    pin = pwinput("Enter your pin: ", mask="*").strip()
    
    workbook = openpyxl.load_workbook(file_name)
    sheet = workbook.active
    
    login_success = False
    user_name = ""
    
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if str(row[0]) == acc_num and str(row[2]) == pin:
            user_name = row[1]
            login_success = True
            break
            
    workbook.close()
    
    if login_success:
        for i in track(range(100), description="[cyan]Verifying..."):
            time.sleep(0.02)
            
        console.print("[bold green]Login successful![/bold green]")
        console.print(f"[green]Welcome back {user_name}[/green]\n")
        
        banking_menu(acc_num)
        return acc_num
    else:
        console.print("[bold red]Invalid account number or PIN[/bold red]")
        return None

#------------------------------------------
# BANK BANNER & MAIN
#------------------------------------------        
def Banner_bank():
    print(emoji.emojize(":bank:"))
    print(emoji.emojize(":money_bag:"))
    print(figlet_format("PY BANK"))
    console.print("[blue] Welcome to the PY BANK [/blue]")
    
    while True:
        print("\n1. Create Account\n2. Login\n3. Exit")
        choice = input("\nEnter your choice : ").strip()
        if choice == "1":
            CREATE_ACCOUNT()
        elif choice == "2":
            login()
        elif choice == "3":
            console.print("[cyan]Thank you for using PY BANK![/cyan]")
            break
        else:
            console.print("[bold red]Enter a correct value (1, 2, or 3)[/bold red]")

if __name__ == "__main__":
    create_excel_file()
    Banner_bank()