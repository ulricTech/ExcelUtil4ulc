import xlwings as xw
filename = r'001.xlsx'

def creatbook():
    app=xw.App(visible=True,add_book=True)
    wb = app.books.add()


 
def test1():
    # book = xw.Book(filename)
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    # 获取 A 列最后一行行数
    lrow = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    search_string = '张三'
    sheet.range('E1').value = 'output'
    output_index = 2
    
    for i in range(2, lrow + 1):
        if search_string in str(sheet.range('A{}'.format(i)).value):
            # temp = str(sheet.range('A{}'.format(i)).value)
            temp = i
            sheet.range('E{}'.format(output_index)).value = temp
            output_index += 1
            break
        

    # book.save()
    # book.close()



