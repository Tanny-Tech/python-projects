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
    


#--------------------banking_menu--------------------------
    def banking_menu(account_no):
        
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
            print("Invalid choice")


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
#---------------cheak_balance-----------------
def check_balance():
    """Displays the current account balance."""
    global balance
    print(f"\nYour current balance is: ₹{balance:.2f}")
#----------------deposit_money------------------
def deposit_money():
    """Allows the user to deposit money into their account."""
    global balance
    try:
        amount = float(input("\nEnter amount to deposit: ₹"))
        if amount > 0:
            balance += amount
            print(f"Successfully deposited ₹{amount:.2f}")
            print(f"New Balance: ₹{balance:.2f}")
        else:
            print("Deposit amount must be greater than zero.")
    except ValueError:
        print("Invalid input! Please enter a valid number.")

def withdraw_money():
    """Allows the user to withdraw money if sufficient balance is available."""
    global balance
    try:
        amount = float(input("\nEnter amount to withdraw: ₹"))
        if amount <= 0:
            print("Withdrawal amount must be greater than zero.")
        elif amount > balance:
            print("Insufficient balance! You cannot withdraw more than your current balance.")
        else:
            balance -= amount
            print(f"Successfully withdrew ₹{amount:.2f}")
            print(f"Remaining Balance: ₹{balance:.2f}")
    except ValueError:
        print("Invalid input! Please enter a valid number.")

def transaction_history():
    """Displays recent transactions (simplified for beginners)."""
    print("\n--- Transaction History ---")
    print("Feature coming soon! (Showing recent activity is up to date)")

# --- Main Program Execution ---
if __name__ == "__main__":
    # Simulating a successful login with a dummy account number
    user_account = "12345678"
    banking_menu(user_account)
#-------------------------banking_menu--------------------------
def banking_menu(account_no):
    # while True:
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
    elif choice == "3":
        pass
    elif choice == "4":
        pass
    elif choice == "5":
        console.print("\n[yellow]Logged out successfully.[/yellow]")
        
    else:
        console.print("[red]Invalid Choice[/red]")
# ----------------------------login------------------------------
def login():
    console.print("[green]---- LOGIN ----[/green]\n")
    acc_num = input("Enter your Account number: ").strip()
    pin = input("Enter your pin: ").strip()
    
    workbook = openpyxl.load_workbook(file_name)
    sheet = workbook.active
    
    login_success = False
    user_name = ""
    
    # Iterate through rows starting from row 2 (skipping header)
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] == acc_num and str(row[2]) == str(pin):
            user_name = row[1]
            login_success = True
            break
            
    workbook.close()
    
    if login_success:
        console.print("[green]Verifying...[/green]")
        for i in track(range(100), description="[cyan]Processing Your Bank Account..."):
            time.sleep(0.02)
            
        console.print("[bold green]Login successful![/bold green]")
        console.print(f"[green]Welcome back {user_name}[/green]\n")
        
        # Proceed to the banking menu
        banking_menu(acc_num)

        return acc_num
    else:
        console.print("[bold red]Invalid account number or PIN[/bold red]")
        return None
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
   



    
        
       
        
            





























# ---------------------------MAINCODE----------------------------
create_excel_file()
Banner_bank()