import openpyxl
from datetime import datetime

# Initialize Excel file with test data
def create_excel_file():
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Login Tests"
    
    # Set the headers
    headers = ["Test ID", "Username", "Password", "Date", "Time of Test", "Name of Tester", "Test Result"]
    sheet.append(headers)
    
    # Add test data (You can adjust usernames and passwords accordingly)
    test_data = [
        [1, "Admin", "admin123", "", "", "Tester 1", ""],
        [2, "WrongUser1", "wrongpass1", "", "", "Tester 1", ""],
        [3, "Admin", "wrongpass2", "", "", "Tester 1", ""],
        [4, "WrongUser2", "admin123", "", "", "Tester 1", ""],
        [5, "Admin", "admin123", "", "", "Tester 1", ""],
    ]
    
    for row in test_data:
        sheet.append(row)
    
    # Save the workbook
    workbook.save("data/login_test_data.xlsx")

create_excel_file()
