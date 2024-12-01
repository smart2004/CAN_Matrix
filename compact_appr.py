import openpyxl
from openpyxl import load_workbook
import re

def load_data(filename):
    wb = openpyxl.load_workbook(filename)
    sheet = wb.active
    return sheet

def extract_first_row_values(sheet):
    first_row_values = []
    for cell in sheet[1][:32]:
        english_first = re.sub("[^\x00-\x7F]+", "", cell.value)
        first_row_values.append(english_first.strip())
    return first_row_values

def extract_row_values(sheet, row_index):
    return [cell.value for cell in sheet[row_index][:32]]

def write_vehicle_data(output_file, vehicle_name, first_row_values, row_values):
    output_file.write(f'Vehicle.{vehicle_name}:\n')
    for i in range(32):
        output_file.write(f'  {first_row_values[i]}: {row_values[i]}\n')

def main():
    sheet = load_data('CAN Matrix.xlsx')
    output_file = open('output.yml', 'w', encoding="utf-8")

    first_row_values = extract_first_row_values(sheet)
    second_row_values = [cell.value for cell in sheet[2][:6]]
    nineth_row_values = [cell.value for cell in sheet[9][:6]]
    fourteenth_row_values = [cell.value for cell in sheet[14][:6]]
    twentythird_row_values = [cell.value for cell in sheet[23][:6]]
    
    for row_index in range(3, 9):  # 3 to 8 corresponds to rows 3 to 8
        row_values = extract_row_values(sheet, row_index)
        write_vehicle_data(output_file, '.'.join(second_row_values), first_row_values, row_values)
        output_file.write('\n')  # Write a newline after each vehicle section
    
    for row_index in range(10, 14):
        row_values = extract_row_values(sheet, row_index)
        write_vehicle_data(output_file, '.'.join(nineth_row_values), first_row_values, row_values)
        output_file.write('\n')

    for row_index in range(15, 23):
        row_values = extract_row_values(sheet, row_index)
        write_vehicle_data(output_file, '.'.join(fourteenth_row_values), first_row_values, row_values)
        output_file.write('\n')
    
    row_values = extract_row_values(sheet, 24)
    write_vehicle_data(output_file, '.'.join(twentythird_row_values), first_row_values, row_values)
    
    output_file.close()


if __name__ == "__main__":
    main()
