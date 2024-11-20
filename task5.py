import openpyxl
file_path_task5 = 'Task5.xlsx'
workbook_task5 = openpyxl.load_workbook(file_path_task5)
sheet_task5 = workbook_task5.active
name = input("Ievadiet vārdu: ")
age = input("Ievadiet vecumu: ")
score = input("Ievadiet rezultātu: ")
sheet_task5.append([name, age, score])
workbook_task5.save(file_path_task5)
txt_file_path = 'task5_data.txt'
with open(txt_file_path, 'a') as txt_file:
    txt_file.write(f"{name}\t{age}\t{score}\n")
print()
