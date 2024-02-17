# I want to read and write data from an Excel file using python
# I want to use the openpyxl library to do this
# write the imports
import openpyxl as xl
import json
import shlex
import tkinter as tk
from tkinter import messagebox

# AE products data
data_headers = ['AE PN', 'Valve', 'Actuator', 'Linkage', 'LD', 'Limit Switch', 'Positioner', 'Linkage',
                'Valve Repair Kit', 'Ormat PN']
data = {}
# AE sale data
sale_headers = ['AE PN', 'Customer', 'Customer PN', 'PO REF', 'From Stock', 'Qty', 'Price', 'Discount', 'Shipment',
                'Final Price']
sales = {}

init = True
src = ''
wb = ''
ps = ''
sd = ''


def query_sales(ae_pn):
    if ae_pn not in sales.keys():
        return f'AE PN {ae_pn} not found'
    customer = input('Enter Customer: ')
    if customer == '':
        return sales[ae_pn]
    # if customer not in sales[ae_pn].values():
    #     print(f'{sales[ae_pn].keys()}')
    #     return f'Customer {customer} not found'
    # result = {f'Sale history of {ae_pn} to {customer}': {}}
    result = {}
    for key, value in sales[ae_pn].items():
        if value['Customer'] == customer:
            result[key] = value
    return result


# load products data
def load_data(ws):
    for row in ws:
        if row[0].row == 1:
            continue
        key = row[0].value
        data[key] = {}
        for cell in row:
            if cell.column == 1:
                continue
            data[key][data_headers[cell.column - 1]] = cell.value


def load_sales(worksheet):
    for key in data.keys():
        sales[key] = {}
    for row in worksheet:
        if row[0].row == 1:
            continue
        key = row[0].value
        ae_ref = row[3].value
        sales[key][ae_ref] = {}
        for cell in row:
            if cell.column == 1:
                continue
            sales[key][ae_ref][sale_headers[cell.column - 1]] = cell.value


def demo():
    global init, src, wb, ps, sd
    # open the workbook
    if init:
        src = input('Please drag file here or enter file path: ')
        # Parse the input using shlex
        parsed_src = shlex.split(src)

        # Join the parsed parts to form the correct file path
        src = " ".join(parsed_src)

        wb = xl.load_workbook(src)
        ps = wb['Data']
        sd = wb['Test']

        load_data(ps)
        init = False

    print('Welcome to the AE Sales History App')
    print('1. View Sales Data')
    print('2. Add Sale Data')
    print('3. View AE Data')
    print('4. Search for AE PN Sale Data')
    print('5. Exit')
    choice = input('Enter your choice: ')
    if choice == '1':
        load_sales(sd)
        pretty = json.dumps(sales, indent=4)
        print(f'Sales Dict: {pretty}')
    elif choice == '2':
        # ws = wb['Test']
        add_sale(sd)
        wb.save(src)
    elif choice == '3':
        load_data(ps)
        pretty = json.dumps(data, indent=4)
        print(f'Data Dict: {pretty}')
    elif choice == '4':
        load_sales(sd)
        ae_pn = int(input('Enter AE PN: '))
        print(f'Sale Data: {json.dumps(query_sales(ae_pn), indent=4)}')
    elif choice == '5':
        exit()
    else:
        print('Invalid choice')
        # demo()
    demo()


# add sale data to excel file
def add_sale(ws):
    ae_pn = input('Enter AE PN: ')
    while int(ae_pn) not in data:
        print('AE PN not found')
        ae_pn = input('Enter AE PN (enter x to cancel):')
        if ae_pn == 'x':
            exit()
    customer = input('Enter Customer: ')
    customer_pn = input('Enter Customer PN: ')
    po_ref = input('Enter PO REF: ')
    from_stock = input('Enter From Stock: ')
    qty = input('Enter Qty: ')
    price = input('Enter Price: ')
    discount = input('Enter Discount: ')
    shipment = input('Enter Shipment: ')
    final_price = f'={qty}*{price}*(1-{int(discount) / 100})+{shipment}'

    ws.append([int(ae_pn), customer, customer_pn, po_ref, from_stock, int(qty), int(price), float(int(discount) / 100),
               int(shipment), final_price])
    # exit()


demo()

# need to initialize the dictionary with the AE PN as the key ( 1 to 58)


# for cell in row:
#     print(cell.value)

# print(cell.value)
# AE PN | Customer | Customer PN | PO REF | From Stock | Qty | Price | Discount | Shipment | Final Price
# ws.append([1,11,111,111,1111,1111])
# wb.save('AE Sales History.xlsx')
