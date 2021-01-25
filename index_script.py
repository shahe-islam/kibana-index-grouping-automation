import openpyxl

#example open_path /Users/Shahe.Islam/developer/ndap-journey/ndap-journey.xlsx

open_path = input("Input the file open path: ")


wb = openpyxl.load_workbook(open_path)
ws = wb['Sheet 1']

###################
##### COLUMNS #####
###################

INDEX_COLUMN = 3
SIZE_COLUMN = 11
HUB_COLUMN = 12

###################
##### SIZES #######
###################

LOG_SIZE_GB = 'gb'
LOG_SIZE_MB = 'mb'

###################
##### JOURNEY #####
###################

journeys = {
    'EDB':['edb', 'containerlogs-edb'],
    'HOME':['home', 'containerlogs-home'],
    'OB':['ob', 'obcompete'],
    'MYNW':['mynw', 'mynationwide'],
    'RAAS':['raas'],
    'NDAP':['ndap', 'ops', 'containerlogs', 'telegraf', 'beat', 'tgw', 'watcher', 'prod', 'monitoring'],
}

for i in range(2, ws.max_row + 1):

    size = ws.cell(row=i, column=9).value

    if LOG_SIZE_GB in size:
        size = size.strip(LOG_SIZE_GB)
        size = str(float(size)*10000)
    elif LOG_SIZE_MB in size:
        size = size.strip(LOG_SIZE_MB)
    else:
        size = 0

    ws.cell(row=i, column=SIZE_COLUMN).value = size

    index = ws.cell(row=i, column=INDEX_COLUMN).value

    for journey, options in journeys.items():
        if any(x in index for x in options): break
    else: journey = 'default'

    ws.cell(row=i, column=HUB_COLUMN).value = journey

#example save_path /Users/Shahe.Islam/developer/ndap-journey/ndap-journey-test2.xlsx

save_path = input("Input the file save path: ")
wb.save(save_path)
print("Your file has been saved at: " + save_path)