import openpyxl
file_name="bank_records_1.xlsx"






#-----excel file creation------
"""This function creates an excel file named bank_records_2"""
def create_excel_file():
    wb=openpyxl.Workbook()
    sheet=wb.active
    sheet.title = "Bank_Records"
    headers=["Account No", "Name", "PIN", "Transaction ID", "Transaction Type", "Amount", "Previous Balance", "Current Balance", "Date-Time"]
    sheet.append(headers)     


    wb.save(file_name)            
    wb.close()                                 

create_excel_file()

                           
                                          


# ----------MAINCODE-------------
