import openpyxl
import os

# Excel file name
excel_file = 'tft_data.xlsx'

# Check if the Excel file exists; if not, create it with headers
if not os.path.exists(excel_file):
    # Create a new workbook and add headers
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "User Data"
    sheet.append(["Name", "Age", "Email"])  # Add headers
    workbook.save(excel_file)

# Function to insert data into the Excel spreadsheet
def insert_data(name, age, email):
    # Load the existing workbook
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    # Append the new data as a row
    sheet.append([name, age, email])
    # Save after appending
    workbook.save(excel_file)
    print("Data saved to Excel successfully!")

# Function to delete the last row in the Excel spreadsheet
def delete_latest_row():
    try:
        # Load the existing workbook
        workbook = openpyxl.load_workbook(excel_file)
        sheet = workbook.active

        # Get the number of rows in the sheet
        max_row = sheet.max_row

        # Ensure there's at least one data row (not just the headers)
        if max_row > 1:
            # Delete the last row
            sheet.delete_rows(max_row)
            workbook.save(excel_file)
            print("Latest row deleted successfully!")
        else:
            print("No data rows to delete.")
    except FileNotFoundError:
        print(f"File '{excel_file}' not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

# Main loop to get user input
def main():
    print("Enter user details to store in the Excel spreadsheet.")
    name = input("Name: ")
    age = input("Age: ")
    email = input("Email: ")

    # Insert data into the Excel spreadsheet
    insert_data(name, age, email)

# Run the main function
if __name__ == '__main__':
    while True:
        user_input = input('What would you like to do? (add/delete/exit) ').lower()
        if 'add' in user_input:
            main()
        elif 'delete' in user_input:
            delete_latest_row()
        elif 'exit' in user_input:
            print("Goodbye!")
            break
        else:
            print("Invalid option. Please type 'add', 'delete', or 'exit'.")