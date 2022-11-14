import xlwings as xw

def createbook():
    app=xw.App(visible=True,add_book=True)
    wb = app.books.add()
    wb.save()