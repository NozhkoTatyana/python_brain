import requests
import json
import openpyxl
import csv
from tkinter import *
from tkinter.filedialog import askdirectory
import os



def write_json(data, filename='_brain.json'):
    with open(filename, 'w') as file:
        json.dump(data, file, indent=2, ensure_ascii=False)




def write_xlsx():
    with open('_brain.json') as file:
        json_data = json.load(file)

        book = openpyxl.Workbook()
        sheet = book.active
        sheet['A1'] = 'NAME'
        sheet['B1'] = 'ARTICUL'
        sheet['C1'] = 'PRODUCT_CODE'
        sheet['D1'] = 'RETAIL_PRICE_UAH'
        sheet['E1'] = 'STOCKS'
        sheet['F1'] = 'AVAILABLE'
        sheet['G1'] = 'BRIEF_DESCRIPTION'
        sheet['H1'] = 'COUNTRY'
        sheet['I1'] = 'MEDIUM_IMAGE'
        sheet['J1'] = 'WARRANTY'

        row = 2

        for li in json_data['result']['list']:
            sheet[row][0].value = li['name']
            sheet[row][1].value = li['articul']
            sheet[row][2].value = li['product_code']
            sheet[row][3].value = li['retail_price_uah']
            #sheet[row][4].value = ' '.join(li['stocks'])
            sheet[row][5].value = ' '.join(li['available'])
            sheet[row][6].value = li['brief_description']
            sheet[row][7].value = li['country']
            sheet[row][8].value = li['medium_image']
            sheet[row][9].value = li['warranty']
            row += 1


    dirname = askdirectory(initialdir=os.getcwd(), title='Please select a directory')
    dirname = book.save(str(dirname) + '/_brain.xlsx')
    book.close()





def get_sid():
    data = {'login': 'test_123', 'password': '6b866754d9c9aa72dd660e5ff5491b2b'}
    r = requests.post('http://api.brain.com.ua/auth', data)
    jsonRes = r.json()
    sid = jsonRes['result']
    return sid



def get_jsoin_product():
    url = path.get() + get_sid()
    r1 = requests.get(url)
    package_json = r1.json()
    return write_json(package_json)



#http://api.brain.com.ua/products/1097/
root = Tk()
root.title('API BRAIN')
root.geometry("300x200")
root.resizable(False, False)
path = StringVar()
path_entry = Entry(root, textvariable=path, background="#ccc", width=45)
path_entry.place(x=10, y=20)
btn = Button(root, text="Send", background="#555", foreground="#ccc", padx="30", pady="1", font="8", command=write_xlsx)
btn.place(x=110, y=110)
root.mainloop()

