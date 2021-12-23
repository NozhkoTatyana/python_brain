# -*- coding: utf-8 -*-
import requests
import json
import os
import win32com.client as win32
import pythoncom
import io





# def write_json_notebook(data, filename='json_notebook'):
#     with open(filename, 'w', encoding='utf8') as file:
#         json.dump(data, file, indent=2, ensure_ascii=False)
#     file.close()
#
#
# def get_sid() -> str:
#     data = {'login': 'test_123', 'password': '6b866754d9c9aa72dd660e5ff5491b2b'}
#     r_sid = requests.post('http://api.brain.com.ua/auth', data)
#     json_data_sid = r_sid.json()
#     sid = json_data_sid['result']
#     return sid
#
#
# def category_notebook():
#     r_category = requests.get('http://api.brain.com.ua/products/1191/' + get_sid())
#     json_data_category_notebook = r_category.json()
#     return json_data_category_notebook
#
#
# def content_notebook():
#     json_data_content = category_notebook()
#     product_list_code = [product_code['product_code'] for product_code in json_data_content['result']['list']]
#     product_list_notebook = []
#     for product_code in product_list_code:
#         r_content = requests.get('http://api.brain.com.ua/product/product_code/' + product_code + '/' + get_sid())
#         json_data = r_content.json()
#         product_list_notebook.append(json_data)
#     write_json_notebook(product_list_notebook)
#         #return product_list_notebook
# content_notebook()


        #data_notebook = json.loads(io.open('json_notebook', encoding='utf-8').read())
        #data_notebook = content_notebook()
def count_option():
    data = json.loads(io.open('json_motherboard', encoding='utf-8').read())
    count = 0
    for i in data[0]['result']['options']:
        count += 1
    return count * 2
count_option()


def record_to_excel():
    data_notebook = json.loads(io.open('json_motherboard', encoding='utf-8').read())
    pythoncom.CoInitializeEx(0)
    ExcelApp = win32.Dispatch('Excel.Application')
    ExcelApp.Visible = True
    wb = ExcelApp.Workbooks.Add()
    ws = wb.Worksheets(1)
    set = count_option()
    header_labels = ['PRODUCT_ID', 'NAME', 'ARTICUL', 'PRODUCT_CODE', 'RETAIL_PRICE_UAH', 'AVAILABLE',
                     'BRIEF_DESCRIPTION', 'COUNTRY', 'MEDIUM_IMAGE', 'DESCRIPTION', 'WARRANTY']

    for i in range(set):
        if i % 2 == 0:
            header_labels.append('VALUE_NAME')
        else:
            header_labels.append('OPTION_NAME')

    rows = []
    for record in data_notebook:
            productID = record['result']['productID']
            name = record['result']['name']
            articul = record['result']['articul']
            product_code = record['result']['product_code']
            retail_price_uah = record['result']['retail_price_uah']
            available = str(record['result']['available'])
            brief_description = record['result']['brief_description']
            country = record['result']['country']
            medium_image = record['result']['medium_image']
            description = record['result']['description']
            warranty = record['result']['warranty']
            indx = 11
            for option in record['result']['options']:
               ws.Cells(2, indx + 1).Value = option ['name']
               row_tracker = 2
               column_size = len(header_labels)
               for row in rows:
                    ws.Range(
                        ws.Cells(row_tracker, 1),
                        ws.Cells(row_tracker, column_size)
                    ).value = row
                    row_tracker += 1


            #options_name = ';'.join(['{0}'.format(option_n['name']) for option_n in record['result']['options']])
                 #options_value = ';'.join(['{0}'.format(option_v['value']) for option_v in record['result']['options']])
            #options_value = [option_v['value'] for option_v in record['result']['options']]

            rows.append([
                     productID, name,
                     articul, product_code,
                     retail_price_uah, available,
                     brief_description, country,
                     medium_image, description, warranty

            ])

    for index, val in enumerate(header_labels):
        ws.Cells(1, index + 1).Value = val
        row_tracker = 2
        column_size = len(header_labels)
        for row in rows:
            ws.Range(
            ws.Cells(row_tracker, 1),
            ws.Cells(row_tracker, column_size)
            ).value = row
            row_tracker += 1
    wb.SaveAs(os.path.join(os.getcwd(), 'notebook.xlsx'))
    ExcelApp.Quit()
    ExcelApp = None
record_to_excel()