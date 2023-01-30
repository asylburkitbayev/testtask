import os
import xlwt

def save_to_excel(path, sheet):
    row_number = 0
    for root, dirs, files in os.walk("."):
        for filename in files:
            file_path = os.path.join(root, filename)
            folder, ext = os.path.splitext(file_path)
            row_number += 1
            sheet.write(row_number, 0, row_number)
            sheet.write(row_number, 1, root)
            sheet.write(row_number, 2, filename)
            sheet.write(row_number, 3, ext)
            print(filename)

def main(path):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('Sheet 1')
    try:
        save_to_excel(path, sheet)
        workbook.save('result.xlsx')
        print('Files saved to result.xlsx')
    except Exception as e:
        print(f'Error: {e}')

if __name__ == '__main__':
    main('testtask')