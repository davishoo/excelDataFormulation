#! python3

import openpyxl, pprint, logging, time
from datetime import datetime

try:
    from hotelDatabase import hotelDB
    print('Import database module: [hotelDatabase] ...')
except:
    print('warning: no hotelDatabase module found!')
    hotelDB = {}



logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')
logging.disable(logging.CRITICAL)
logging.debug('Start of program')

print('Opening workbook...')
wb = openpyxl.load_workbook('inputHotelAttribute.xlsx')
sheet = wb['Sheet1']

hotelSubDict = {}
count = 0

# Fill in hotel data from input file
print('Reading rows...')
for row in range(2, sheet.max_row + 1):
    # read excel input data
    hotelName = sheet['A' + str(row)].value
    hotelLocation = sheet['B' + str(row)].value
    address = sheet['C' + str(row)].value
    hotelSales = sheet['D' + str(row)].value
    mobilePhone = sheet['E' + str(row)].value
    officePhone = sheet['F' + str(row)].value
    wechat = sheet['G' + str(row)].value
    email = sheet['H' + str(row)].value


    # Make sure the key for this hotel exists.
    hotelDB.setdefault(hotelName, [])

    # set default value
    hotelSubDict.setdefault('hotelLocation',  'N/A')
    hotelSubDict.setdefault('address',  'N/A')
    hotelSubDict.setdefault('hotelSales',  'N/A')
    hotelSubDict.setdefault('mobilePhone',  'N/A')
    hotelSubDict.setdefault('officePhone',  'N/A')
    hotelSubDict.setdefault('wechat',  'N/A')
    hotelSubDict.setdefault('email',  'N/A')

    logging.debug('1. hotelSubDict = %s ' %hotelSubDict)

    #fill in hotel data
    hotelSubDict['hotelLocation'] = hotelLocation
    hotelSubDict['address'] = address
    hotelSubDict['hotelSales'] = hotelSales
    hotelSubDict['mobilePhone'] = mobilePhone
    hotelSubDict['officePhone'] = officePhone
    hotelSubDict['wechat'] = wechat
    hotelSubDict['email'] = email
    logging.debug('2. hotelSubDict = %s ' %hotelSubDict)  #check if data has been assigned successfully

    HotelDataValue = hotelDB[hotelName]           #data type: list
    logging.debug('3 hotelDB[%s] = %s ' %(hotelName, HotelDataValue))

    if  len(HotelDataValue) > 0:
        print('Hotel data already exists for [%s]\n' %(hotelName))

        while True:
            decision = input('1. Skip this input data? \n' + '2. Append this input data?\n' + 'Enter the choice number: ')
            if decision in [ '1' , '2']:
                break               
            else:
                print('Wrong input, pls select again...')
                continue

        if decision == '1':
            print('Data skipped...')
            continue
        else:
            print('Data appended...')

    HotelDataValue += [hotelSubDict]
    count += 1

logging.debug(' End of loop...') 

timeStamp = time.strftime("%Y-%m-%d %H:%M",time.localtime(time.time()))

print('Writing results...')
resultFile = open('hotelDatabase.py', 'w', encoding = "utf-8")
resultFile.write('import datetime\n\n')
resultFile.write('# Update time: %s\n\n'%timeStamp)
resultFile.write('hotelDB = ' + pprint.pformat(hotelDB))
resultFile.close()

print('Done! total %s data are write into the database.  \n'%count)

    
