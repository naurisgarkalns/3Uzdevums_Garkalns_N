from openpyxl import load_workbook
file_path_task3 = "Task3.xlsx"
output_path_task3 = "uppercase_task3.txt"
workbook_task3 = load_workbook(file_path_task3)
sheet_task3 = workbook_task3.active
uppercase_texts = [str(cell.value).upper() for cell in sheet_task3['A'][1:] if cell.value is not None]
with open(output_path_task3, 'w') as file:
    for text in uppercase_texts:
        file.write(text + "\n")
len(uppercase_texts)