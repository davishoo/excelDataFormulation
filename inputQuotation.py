#! python3

import openpyxl, pprint, logging
from datetime import datetime


logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')

logging.disable(logging.CRITICAL)
logging.debug('Start of program')

print('Opening workbook...')
wb = openpyxl.load_workbook('quotationData.xlsx')
sheet = wb['Sheet1']

quotationData = {}
QuoteDateValueDict = {}

# Fill in quotation data from input file
print('Reading rows...')
for row in range(3, sheet.max_row + 1):
    # read excel input data
    Project  = sheet['A' + str(row)].value
    Hotel = sheet['B' + str(row)].value
    RoomType = sheet['C' + str(row)].value
    QuoteDate = sheet['D' + str(row)].value     # need change data type
    ExpireDate = sheet['E' + str(row)].value
    Day1 = sheet['F' + str(row)].value
    Day2 = sheet['G' + str(row)].value
    Day3 = sheet['H' + str(row)].value
    Day4 = sheet['I' + str(row)].value
    Tax = sheet['J' + str(row)].value
    Breakfast = sheet['K' + str(row)].value
    breakfastPrice = sheet['L' + str(row)].value
    Prepayment = sheet['M' + str(row)].value
    minRoomQty = sheet['N' + str(row)].value
    QuotePerson = sheet['O' + str(row)].value

    # Make sure the key for this Project exists.
    quotationData.setdefault(Project, {})
    # Make sure the key for this Hotel in this state exists.
    quotationData[Project].setdefault(Hotel, {})
    # Make sure the key for this room type in this state exists.
    quotationData[Project][Hotel].setdefault(RoomType, {})
     # Make sure the key for this quote date in this state exists.
    quotationData[Project][Hotel][RoomType].setdefault(QuoteDate, [])


    #set default value
    QuoteDateValueDict.setdefault('ExpireDate', 'N/A')
    QuoteDateValueDict.setdefault('2019-6-19', 'N/A')
    QuoteDateValueDict.setdefault('2019-6-20', 'N/A')
    QuoteDateValueDict.setdefault('2019-6-21', 'N/A')
    QuoteDateValueDict.setdefault('2019-6-22', 'N/A')
    QuoteDateValueDict.setdefault('Tax', 'N/A')
    QuoteDateValueDict.setdefault('Breakfast', 'N/A')
    QuoteDateValueDict.setdefault('breakfastPrice', 'N/A')
    QuoteDateValueDict.setdefault('Prepayment', 'N/A')
    QuoteDateValueDict.setdefault('minRoomQty', 'N/A')
    QuoteDateValueDict.setdefault('QuotePerson', 'N/A')

    logging.debug('1. QuoteDateValueDict = %s ' %QuoteDateValueDict)

    # fill in quotation data 
    QuoteDateValueDict['ExpireDate'] = ExpireDate
    QuoteDateValueDict['2019-6-19'] = Day1
    QuoteDateValueDict['2019-6-20'] = Day2
    QuoteDateValueDict['2019-6-21'] = Day3
    QuoteDateValueDict['2019-6-22'] = Day4
    QuoteDateValueDict['Tax'] = Tax
    QuoteDateValueDict['Breakfast'] = Breakfast
    QuoteDateValueDict['breakfastPrice'] = breakfastPrice
    QuoteDateValueDict['Prepayment'] = Prepayment
    QuoteDateValueDict['minRoomQty'] = minRoomQty
    QuoteDateValueDict['QuotePerson'] = QuotePerson

    logging.debug('2. QuoteDateValueDict = %s ' %QuoteDateValueDict)  #check if data has been assigned successfully

   
    fullQuotationDict = quotationData[Project][Hotel][RoomType][QuoteDate]
    logging.debug('3. quotationData[Project][Hotel][RoomType][QuoteDate] = %s ' %fullQuotationDict)

    if  len(fullQuotationDict) > 0:
        print('Quotation data already exists for %s-%s-%s' %(Project, Hotel, RoomType) + ' : ' + QuoteDate.strftime('%Y-%m-%d'))

        while True:
            decision = input('1. Skip this input data? \n' + '2. Append this input data?\n' + 'Enter the choice number: ')
            if not decision in [ '1' , '2']:
                print('Wrong input, pls select again...')
                continue
            else:
                break
            
        if decision == '1':
            print('Data skipped...')
            continue
        else:
            print('Data appended...')
            
    fullQuotationDict += [QuoteDateValueDict] 

logging.debug(' End of loop...') 
   
# Open a new text file and write the contents of countyData to it.
print('Writing results...')
resultFile = open('hotelQuotation.py', 'w', encoding = "utf-8")
resultFile.write('import datetime\n')
resultFile.write('quotationData = ' + pprint.pformat(quotationData))
resultFile.close()
print('Done.')
