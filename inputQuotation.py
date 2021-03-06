#! python3

import openpyxl, pprint, logging, time
from datetime import datetime

try:
    from hotelQuotation import quotationData
    print('Import database module: [hotelQuotation] ...')
except:
    print('warning: no hotelQuotation module found!')
    quotationData = {}

logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')
logging.disable(logging.CRITICAL)
logging.debug('Start of program')

print('Opening workbook...')
wb = openpyxl.load_workbook('quotationData.xlsx')
sheet = wb['Sheet1']

QuoteDateValueDict = {}
count = 0

# Fill in quotation data from input file
print('Reading rows...\n\n')
for row in range(3, sheet.max_row + 1):
    # read excel input data
    Project  = sheet['A' + str(row)].value
    quotationSource = sheet['B' + str(row)].value          #insert
    Hotel = sheet['C' + str(row)].value
    RoomType = sheet['D' + str(row)].value
    QuoteDate = sheet['E' + str(row)].value     
    ExpireDate = sheet['F' + str(row)].value
    Day1 = sheet['G' + str(row)].value
    Day2 = sheet['H' + str(row)].value
    Day3 = sheet['I' + str(row)].value
    Day4 = sheet['J' + str(row)].value
    Tax = sheet['K' + str(row)].value
    Breakfast = sheet['L' + str(row)].value
    breakfastPrice = sheet['M' + str(row)].value
    Prepayment = sheet['N' + str(row)].value
    minRoomQty = sheet['O' + str(row)].value
    QuotePerson = sheet['P' + str(row)].value

    # Make sure the key for this Project exists.
    quotationData.setdefault(Project, {})
    # Make sure the key for this Hotel in this state exists.
    quotationData[Project].setdefault(quotationSource, {})    #insert
    # Make sure the key for this Hotel in this state exists.
    quotationData[Project][quotationSource].setdefault(Hotel, {})
    # Make sure the key for this room type in this state exists.
    quotationData[Project][quotationSource][Hotel].setdefault(RoomType, {})
     # Make sure the key for this quote date in this state exists.
    quotationData[Project][quotationSource][Hotel][RoomType].setdefault(QuoteDate, [])


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

   
    fullQuotationValue = quotationData[Project][quotationSource][Hotel][RoomType][QuoteDate]                #data type : list
    logging.debug('3. quotationData[%s][%s][%s][%s][%s] = %s ' %(Project, quotationSource,
                                                                 Hotel, RoomType, QuoteDate.strftime('%Y-%m-%d'), fullQuotationValue) ) 

    if  len(fullQuotationValue) > 0:
        print('Quotation data already exists for [%s]-[%s]-[%s]-[%s]-[%s]\n' %(Project,
                                                                               quotationSource, Hotel, RoomType, QuoteDate.strftime('%Y-%m-%d')))

        while True:
            decision = input('1. Skip this input data? \n' + '2. Append this input data?\n' + 'Enter the choice number: ')
            if decision in [ '1' , '2']:             
                break
            else:
                print('\nWrong input, pls select again...\n')
                continue
            
        if decision == '1':
            print('Data skipped...\n')
            continue
        else:
            print('Data appended...\n')
            
    fullQuotationValue += [QuoteDateValueDict]
    count += 1

logging.debug(' End of loop...') 
   

timeStamp = time.strftime("%Y-%m-%d %H:%M",time.localtime(time.time()))

print('Writing results...\n')
resultFile = open('hotelQuotation.py', 'w', encoding = "utf-8")
resultFile.write('import datetime\n\n')
resultFile.write('# Update time: %s\n\n'%timeStamp)
resultFile.write('quotationData = ' + pprint.pformat(quotationData))
resultFile.close()

print('Done! total %s data are write into the database.  \n'%count)
