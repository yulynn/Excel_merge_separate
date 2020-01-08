from pathlib import Path
import xlwings as xw
from xlwings.constants import FileFormat

root_path = Path.cwd()
first = 0
file_name = root_path / '汇总.xlsx'
new_books = []
for item in root_path.glob('汇总_*.xls*'):
    item.unlink()
with open('process.log', 'w') as f:
    for dir_item in root_path.iterdir():
        if (dir_item.is_file()):
            if (dir_item.suffix.lower() == '.xls'
                    or dir_item.suffix.lower() == '.xlsx'):
                if (first == 0):
                    wb = xw.Book(str(dir_item))
                    first = 1
                    ws = wb.sheets
                    sht_nums = len(ws)
                    for i in range(sht_nums):
                        ws = wb.sheets[i]
                        new_wb = xw.Book()
                        new_sht = new_wb.sheets[0]
                        ws.api.Copy(Before=new_sht.api)
                        new_books.append(new_wb)
                    f.write('{} files will be exported\n'.format(sht_nums))
                    f.write('{} is processed successfully\n'.format(
                        str(dir_item)))
                else:
                    wb1 = xw.Book(str(dir_item))
                    if (len(wb1.sheets) != sht_nums):
                        f.write(
                            'Error! {} is not processed, the workbook has {} sheets which is different from initial num {}\n'
                            .format(str(dir_item), len(wb1.sheets), sht_nums))
                        wb1.close()
                        continue
                    for i in range(sht_nums):
                        new_wb = new_books[i]
                        ws1 = wb1.sheets[i]
                        ws1.api.Copy(Before=new_wb.sheets[0].api)
                    f.write('{} is processed successfully\n'.format(
                        str(dir_item)))
                    wb1.close()
                # wb1.app.quit()
for i in range(sht_nums):
    new_wb = new_books[i]
    new_wb.sheets[-1].delete()
    new_wb.api.SaveAs(
        str(root_path) + r'/汇总_' + str(i) + '.xls', FileFormat.xlExcel8)
#wb.close()
print('done')
wb.app.quit()
