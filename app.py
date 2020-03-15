import xlrd
import os
import shutil
import time

if not os.path.isdir('./dataset'):
    os.mkdir('./dataset')

workbook = xlrd.open_workbook('humir.xlsx')

for i in range(workbook.nsheets):
    woorksheet = workbook.sheet_by_index(i)

    bookname = woorksheet.name
    if os.path.isdir('./dataset/' + bookname):
        shutil.rmtree('./dataset/' + bookname)
    
    os.mkdir('./dataset/' + bookname)
    os.mkdir('./dataset/' + bookname + '/positive')
    os.mkdir('./dataset/' + bookname + '/negative')

    positiveIndex = 0
    negativeIndex = 0
    print("Kitap Okunuyor!!!\nİşlem süresi:" + str(woorksheet.nrows * 0.2 / 60) + "dk.")

    for x in range(woorksheet.nrows):
        if woorksheet.cell(x,1).value == 'Positive':
            file = open("./dataset/" + bookname + "/positive/" +"row" + str(positiveIndex) + ".txt", 'w')
            file.write(woorksheet.cell(x,0).value)
            time.sleep(0.2)
            positiveIndex = positiveIndex + 1
        else:
            file = open("./dataset/" + bookname + "/negative/" +"row" + str(negativeIndex) + ".txt", 'w')
            file.write(woorksheet.cell(x,0).value)
            time.sleep(0.2)
            negativeIndex = negativeIndex + 1
