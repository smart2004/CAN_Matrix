import openpyxl
from openpyxl import load_workbook
import re

wb = openpyxl.load_workbook('CAN Matrix.xlsx')
sheet = wb.active
output_file = open('output.yml', 'w', encoding="utf-8")

first_row_values = []
for cell in sheet[1][:32]:
    english_first = re.sub("[^\x00-\x7F]+", "", cell.value)
    first_row_values.append(english_first.strip())

second_row_values = [cell.value for cell in sheet[2][:6]]
third_row_values = [cell.value for cell in sheet[3][:32]]
fourth_row_values = [cell.value for cell in sheet[4][:32]]
fifth_row_values = [cell.value for cell in sheet[5][:32]]
sixth_row_values = [cell.value for cell in sheet[6][:32]]
seventh_row_values = [cell.value for cell in sheet[7][:32]]
eights_row_values = [cell.value for cell in sheet[8][:32]]

output_file.write(f'Vehicle.{'.'.join(second_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {third_row_values[i]}\n'
        )
    
output_file.write('\n' f'Vehicle.{'.'.join(second_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {fourth_row_values[i]}\n'
        )

output_file.write('\n' f'Vehicle.{'.'.join(second_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {fifth_row_values[i]}\n'
        )
    
output_file.write('\n' f'Vehicle.{'.'.join(second_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {sixth_row_values[i]}\n'
        )

output_file.write('\n' f'Vehicle.{'.'.join(second_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {seventh_row_values[i]}\n'
        )

output_file.write('\n' f'Vehicle.{'.'.join(second_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {eights_row_values[i]}\n'
        )

nineth_row_values = [cell.value for cell in sheet[9][:6]]
tenth_row_values = [cell.value for cell in sheet[10][:32]]
eleventh_row_values = [cell.value for cell in sheet[11][:32]]
twelveth_row_values = [cell.value for cell in sheet[12][:32]]
thirteenth_row_values = [cell.value for cell in sheet[13][:32]]

output_file.write('\n' f'Vehicle.{'.'.join(nineth_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {tenth_row_values[i]}\n'
        )

output_file.write('\n' f'Vehicle.{'.'.join(nineth_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {eleventh_row_values[i]}\n'
        )

output_file.write('\n' f'Vehicle.{'.'.join(nineth_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {twelveth_row_values[i]}\n'
        )

output_file.write('\n' f'Vehicle.{'.'.join(nineth_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {thirteenth_row_values[i]}\n'
        )

fourteenth_row_values = [cell.value for cell in sheet[14][:6]]
fifteenth_row_values = [cell.value for cell in sheet[15][:32]]
sixteenth_row_values = [cell.value for cell in sheet[16][:32]]
seventeenth_row_values = [cell.value for cell in sheet[17][:32]]
eighteenth_row_values = [cell.value for cell in sheet[18][:32]]
nineteenth_row_values = [cell.value for cell in sheet[19][:32]]
twentyth_row_values = [cell.value for cell in sheet[20][:32]]
twentyfirst_row_values = [cell.value for cell in sheet[21][:32]]
twentysecond_row_values = [cell.value for cell in sheet[22][:32]]

output_file.write('\n' f'Vehicle.{'.'.join(fourteenth_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {fifteenth_row_values[i]}\n'
        )
    
output_file.write('\n' f'Vehicle.{'.'.join(fourteenth_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {sixteenth_row_values[i]}\n'
        )
    
output_file.write('\n' f'Vehicle.{'.'.join(fourteenth_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {seventeenth_row_values[i]}\n'
        )
    
output_file.write('\n' f'Vehicle.{'.'.join(fourteenth_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {eighteenth_row_values[i]}\n'
        )

output_file.write('\n' f'Vehicle.{'.'.join(fourteenth_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {nineteenth_row_values[i]}\n'
        )
    
output_file.write('\n' f'Vehicle.{'.'.join(fourteenth_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {twentyth_row_values[i]}\n'
    )

output_file.write('\n' f'Vehicle.{'.'.join(fourteenth_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {twentyfirst_row_values[i]}\n'
    )

output_file.write('\n' f'Vehicle.{'.'.join(fourteenth_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {twentysecond_row_values[i]}\n'
    )

twentythird_row_values = [cell.value for cell in sheet[23][:6]]
twentyfourth_row_values = [cell.value for cell in sheet[24][:32]]

output_file.write('\n' f'Vehicle.{'.'.join(twentythird_row_values)}:\n')

for i in range(32):
    output_file.write(
        f'  {first_row_values[i]}: {twentyfourth_row_values[i]}\n'
    )

output_file.close()
