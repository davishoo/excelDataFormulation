import openpyxl, pprint, logging, time
from openpyxl.styles import Font
from datetime import datetime

logging.disable(logging.CRITICAL)
logging.debug('Start of program')

try:
    from hotelQuotation import quotationData
    print('Import database module: [hotelQuotation] ...')
except  ModuleNotFoundError:
    print('warning: no hotelQuotation module found!')

#Add more attributes to database



wb = openpyxl.Workbook()
sheet = wb['Sheet']


#fill out the head of the sheet


sheet['A1'] = 'hello'
wb.save('hotel_quotation.xlsx')






        
