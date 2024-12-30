import xlwings as xw
import requests
import os
from datetime import datetime, timezone

# If config page is http://localhost:38080/page/index.html, server address is:
DLS_SERVER = 'http://localhost:38080/' 
# license item ids
LICENSE_ITEMS = ['123456','234567']
# can be found in F12 Console -> Network -> Fetch/XHR ->
# http://localhost:38080/license/item/ -> Request Headers
DynamsoftLTSTokenV2 = '?????' 

book_absolute_path = os.path.realpath(__file__+'/../watch-license.xlsx')
utcnow = datetime.now(timezone.utc)
strutcnow = str(utcnow.year)+'.'+str(utcnow.month).zfill(2)+'.'+str(utcnow.day).zfill(2)

sheet_name_prefix = str(utcnow.year) + 's' + str((utcnow.month+2)//3)
sheet_name_suffix:str = None # empty or a-z, should enough
sheet_name:str = None

with xw.App() as app:
    book: xw.Book = None
    if os.path.exists(book_absolute_path):
        book = app.books.open(book_absolute_path)
    else:
        book = app.books[0]

    # get last sheet of this season
    candidate_sheet_name = [s for s in book.sheet_names if s.startswith(sheet_name_prefix)]
    if len(candidate_sheet_name):
        sheet_name = max(candidate_sheet_name)
    sheet: xw.Sheet = None
    if sheet_name:
        sheet_name_suffix = sheet_name[len(sheet_name_prefix):]
        sheet = book.sheets[sheet_name]
    else:
        sheet_name = sheet_name_prefix
        sheet = book.sheets.add(sheet_name)
    

    if sheet['B1'].value:

        # check license items
        is_license_items_different = False
        for i in range(len(LICENSE_ITEMS)):
            license_item = LICENSE_ITEMS[i]
            if 'item: '+license_item != sheet['A'+str(6+2*i)].value:
                is_license_items_different = True
                break
        
        # if different, need create new sheet
        if is_license_items_different:
            if not sheet_name_suffix:
                sheet_name_suffix = 'a'
            else:
                if sheet_name_suffix >= 'z':
                    raise Exception('The `sheet_name_suffix` is already `z`. No more sheet can be added.')
                sheet_name_suffix = chr(ord(sheet_name_suffix)+1)
            sheet_name = sheet_name_prefix + sheet_name_suffix
            sheet = book.sheets.add(sheet_name)
    
    if not sheet['B1'].value:
        # need to initialize the header
        sheet['A1'].column_width = 12
        sheet['B1'].column_width = 25

        sheet['B1'].value = 'Date:'
        sheet['B2'].value = 'Total used:'
        sheet['B3'].value = 'Total available:'
        sheet['B4'].value = 'Unused licenses:'
        sheet['B5'].value = 'Change from previous:'

        for i in range(len(LICENSE_ITEMS)):
            sheet['A'+str(6+2*i)].value = 'item: '+LICENSE_ITEMS[i]
            sheet['B'+str(6+2*i)].value = 'License '+str(i+1)+' used'
            sheet['B'+str(7+2*i)].value = 'License '+str(i+1)+' available'

    column = 2
    # find an empty column
    while sheet[0,column].value:
        column = column + 1
    else:
        sheet[0,column].column_width = 10
        sheet[0,column].value = strutcnow
        total_used = 0
        total_available = 0

        # data for every license item
        for i in range(len(LICENSE_ITEMS)):
            rep = requests.get(DLS_SERVER+'license/item/'+LICENSE_ITEMS[i], headers = {
                'DynamsoftLTSTokenV2': DynamsoftLTSTokenV2
            })
            license_item = rep.json()
            total_used = total_used + license_item['usedCount']
            sheet[5+2*i, column].value = license_item['usedCount']
            total_available = total_available + license_item['quantity']
            sheet[6+2*i, column].value = license_item['quantity']
        
        # Statistics for this day
        sheet[1, column].value = total_used
        sheet[2, column].value = total_available
        sheet[3, column].value = total_available - total_used
        if column > 2 and sheet[3, column - 1].value:
            sheet[4, column].value = sheet[3, column].value - sheet[3, column - 1].value

    
    book.save(book_absolute_path)
