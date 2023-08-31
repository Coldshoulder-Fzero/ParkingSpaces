import os
import xlrd
import subprocess


def search_and_run_program(folder_path):
    # Search for Excel files in the specified folder
    excel_files = [file for file in os.listdir(folder_path) if file.endswith('.xls')]


    if excel_files:
        with open('GarageOccupancy.txt', 'a') as file:
            for excel_file in excel_files:
                excel_file_path = os.path.join(folder_path, excel_file)
                nuparks(excel_file_path, file)
            # Delete the Excel file after the program executes
                os.remove(excel_file_path)
            # Open the text file with the default text editor
        subprocess.run(['notepad.exe', 'GarageOccupancy.txt'])

        os.remove('GarageOccupancy.txt')
    else:
        print("No Excel files found in the folder.")



def nuparks(excel_file_path, file):
    print("hello")
    workbook = xlrd.open_workbook(excel_file_path)
    sheet = workbook.sheet_by_index(0)  # Assuming you want the first sheet

    PG1ST = 586
    PG2ST = 580
    PG3ST = 1202
    PG4ST = 995
    PG5ST = 1611
    PG6ST = 1637

    last_row = sheet.nrows - 1
    # Extract date from column C
    date_value = sheet.cell_value(last_row, 2)  # Assuming column C is index 2
    formatted_date = xlrd.xldate.xldate_as_datetime(date_value, workbook.datemode)

    # Extract the time from the formatted date and format it as "h:mm am/pm"
    formatted_time = formatted_date.strftime("%I:%M%p")  # Use %I to format hour without leading zeros

    # Remove leading zero from the hour if present
    if formatted_time.startswith("0"):
        formatted_time = formatted_time[1:]

    # Extract student parking numbers and perform subtraction

    student_parking_pg1 = sheet.cell_value(last_row, 4)
    student_parking_pg2 = sheet.cell_value(last_row, 9)
    student_parking_pg3 = sheet.cell_value(last_row, 12)
    student_parking_pg4 = sheet.cell_value(last_row, 16)
    student_parking_pg5 = sheet.cell_value(last_row, 19)
    student_parking_pg6 = sheet.cell_value(last_row, 23)

    """
    print(student_parking_pg1)
    print(student_parking_pg2)
    print(student_parking_pg3)
    print(student_parking_pg4)
    print(student_parking_pg5)
    print(student_parking_pg6)
    """

    PG1studentParkingLeft = PG1ST - student_parking_pg1
    PG2studentParkingLeft = PG2ST - student_parking_pg2
    PG3studentParkingLeft = PG3ST - student_parking_pg3
    PG4studentParkingLeft = PG4ST - student_parking_pg4
    PG5studentParkingLeft = PG5ST - student_parking_pg5
    PG6studentParkingLeft = PG6ST - student_parking_pg6

    parkingGarages = [PG1studentParkingLeft, PG2studentParkingLeft, PG3studentParkingLeft, PG4studentParkingLeft,
                      PG5studentParkingLeft, PG6studentParkingLeft]
    # Print the results

    # Replace negative values with "Full"
    sCountTotal = 0
    for i in range(len(parkingGarages)):
        if parkingGarages[i] > 0:
            sCountTotal = parkingGarages[i] + sCountTotal
        else:
            parkingGarages[i] = "Full"

    # Close the Excel file (not necessary for xlrd)
    # workbook.close()

    PG1FT = 342
    PG2FT = 342
    PG3FT = 145
    PG4FT = 394
    PG5FT = 195
    PG6FT = 197

    # Extract student parking numbers and perform subtraction

    f_parking_pg1 = sheet.cell_value(last_row, 6)
    f_parking_pg2 = sheet.cell_value(last_row, 10)
    f_parking_pg3 = sheet.cell_value(last_row, 13)
    f_parking_pg4 = sheet.cell_value(last_row, 17)
    f_parking_pg5 = sheet.cell_value(last_row, 20)
    f_parking_pg6 = sheet.cell_value(last_row, 24)

    """
    print(f_parking_pg1)
    print(f_parking_pg2)
    print(f_parking_pg3)
    print(f_parking_pg4)
    print(f_parking_pg5)
    print(f_parking_pg6)
    """
    PG1fParkingLeft = PG1FT - f_parking_pg1
    PG2fParkingLeft = PG2FT - f_parking_pg2
    PG3fParkingLeft = PG3FT - f_parking_pg3
    PG4fParkingLeft = PG4FT - f_parking_pg4
    PG5fParkingLeft = PG5FT - f_parking_pg5
    PG6fParkingLeft = PG6FT - f_parking_pg6

    fparkingGarages = [PG1fParkingLeft, PG2fParkingLeft, PG3fParkingLeft, PG4fParkingLeft, PG5fParkingLeft,
                       PG6fParkingLeft]
    # Print the results

    # Replace negative values with "Full"
    fCountTotal = 0
    for i in range(len(fparkingGarages)):

        if fparkingGarages[i] > 0:
            fCountTotal = fparkingGarages[i] + fCountTotal
        elif fparkingGarages[i] < 0:
            fparkingGarages[i] = "Full"


    file.write("\n")
    file.write(f"For the {formatted_time} time frame, the following student spaces were available at MMC\n")
    file.write(f"Pg1 - {parkingGarages[0]}\n")
    file.write(f"Pg2 - {parkingGarages[1]}\n")
    file.write(f"Pg3 - {parkingGarages[2]}\n")
    file.write(f"Pg4 - {parkingGarages[3]}\n")
    file.write(f"Pg5 - {parkingGarages[4]}\n")
    file.write(f"Pg6 - {parkingGarages[5]}\n")
    file.write("Lot1 - \n")  # Replace with actual Lot1 value
    file.write("Lot3 - \n")  # Replace with actual Lot3 value
    file.write("Lot5 - \n")  # Replace with actual Lot5 value
    file.write("Lot7 - \n")  # Replace with actual Lot7 value
    file.write("Lot9 - \n")  # Replace with actual Lot9 value
    file.write(f"Total numbers of available student spaces at the MMC: {sCountTotal}\n")

    file.write(
        f"\nFor the {formatted_time} time frame, the following admin/faculty/staff spaces are available at MMC:\n")
    file.write(f"Pg1 - {fparkingGarages[0]}\n")
    file.write(f"Pg2 - {fparkingGarages[1]}\n")
    file.write(f"Pg3 - {fparkingGarages[2]}\n")
    file.write(f"Pg4 - {fparkingGarages[3]}\n")
    file.write(f"Pg5 - {fparkingGarages[4]}\n")
    file.write(f"Pg6 - {fparkingGarages[5]}\n")
    file.write("Lot1 - \n")  # Replace with actual Lot1 value
    file.write("Lot3 - \n")  # Replace with actual Lot3 value
    file.write("Lot5 - \n")  # Replace with actual Lot5 value
    file.write("Lot7 - \n")  # Replace with actual Lot7 value
    file.write("Lot9 - \n")  # Replace with actual Lot9 value
    file.write(
        f"Total number of available admin/faculty/staff spaces in the areas posted above at MMC: {fCountTotal}\n")

    """file.write("(The following report was created by Joseph's Python program be wary of any outliers)\n\n")"""

    """if fparkingGarages[4] > 160:
        file.write("The faculty count for pg5 may be incorrect. Please verify this report.\n")
    if fparkingGarages[5] > 160:
        file.write("The faculty count for pg6 may be incorrect. Please verify this report.\n")"""

    file.write("\n\n")


if __name__ == "__main__":
    shared_folder = r'C:\Users\jgabriem\Desktop\Garage excel'
    search_and_run_program(shared_folder)

