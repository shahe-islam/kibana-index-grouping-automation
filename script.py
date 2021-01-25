import openpyxl

directory = "/Users/Shahe.Islam/developer/ndap-journey/"
filename = "ndap-journey.xlsx"

wb = openpyxl.load_workbook(directory + filename)
ws = wb['Sheet 1']

###################
##### SIZES #######
###################

GB = 'gb'
MB = 'mb'

###################
##### JOURNEY #####
###################

EDB = ['edb', 'containerlogs-edb']
HOME = ['home', 'containerlogs-home']
OB = ['ob', 'obcompete']
MYNW = ['mynw', 'mynationwide']
RAAS = ['raas']
NDAP = ['ndap', 'ops', 'containerlogs', 'telegraf', 'beat', 'tgw', 'watcher', 'prod', 'monitoring']

for i in range(2, ws.max_row + 1):

    size = ws.cell(row=i, column=9).value

    if GB in size:
        size = size.strip(GB)
        size = str(float(size)*10000)
    elif MB in size:
        size = size.strip(MB)
    else:
        size = 0

    ws.cell(row=i, column=11).value = size

    index = ws.cell(row=i, column=3).value

    if any(x in index for x in EDB):
        journey = 'EDB'
    elif any(x in index for x in HOME):
        journey = 'HOME'
    elif any(x in index for x in OB):
        journey = 'OB'
    elif any(x in index for x in MYNW):
        journey = 'MYNW'
    elif any(x in index for x in RAAS):
        journey = 'RAAS'
    elif any(x in index for x in NDAP):
        journey = 'NDAP'
    else:
        journey = 'default'
    
    ws.cell(row=i, column=12).value = journey
    journey = ''

test_filename = "ndap-journey-test.xlsx"
wb.save(directory + test_filename)