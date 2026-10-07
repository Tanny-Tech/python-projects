import openpyxl
import os





from rich.console import Console
from pyfiglet import figlet_format


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
    print(figlet_format("PY BANK"))
    console.print("[blue] Welcome to the PY BANK [/blue]")
    print("1. Create Account\n2. Login\n3. Exit")
    choice=int(input("\n\nEnter your choice :"))
    while True :
            
        if choice==1:
            pass
        elif choice==2:
            pass
        elif choice==3:
            break

        else:
            Console.print("[red]Enter correct value")



    
 # ----------MAINCODE-------------

                                        
Banner_bank()

