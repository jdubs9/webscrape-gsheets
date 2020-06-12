from time import sleep

import gspread
from oauth2client.service_account import ServiceAccountCredentials

from formatting import format_sheet
from webscrape import webscrape_df

SHEET_NAME = "Iron Ore Open Interests"
SHEET_DATE = "18 May"
SHEET_EMAILS = ['copy client_email from the client_secret.json here', 'own email']
URL='https://api2.sgx.com/sites/default/files/reports/settlement-prices/2020/05/wcm%40sgx_en%40iron_dsp%4018-May-2020%40Iron_Ore_Options_DSP.html'
NUM_COL = 12

df = webscrape_df(URL, NUM_COL)

#take out only columns needed
df = df[['Commodity','Contract Year','Contract Month', 'Contract Type', 'Strike', 'Open Interest']]

df2=df.copy()
df2=df2[['Contract Year','Contract Month']]
df2=df2.drop_duplicates() #get unique month, year pairs

ndict = {}

unique_dates =[]
second_row = []
for index, row in df2.iterrows(): #for the top 2 rows in google sheets
    date = str(row['Contract Month']) + "/" + str(row['Contract Year'])
    unique_dates.extend([date, date, date, date])
    second_row.extend(['Strike', 'Call OI', 'Put OI', 'Total'])
    ndict[date] = []


dates = []
strikes_commodity_list = []
for index, row in df.iterrows(): #editing dataframe
    date = str(row['Contract Month']) + "/" + str(row['Contract Year'])
    dates.append(date)
    strike_commodity = str(row['Strike']) + " (" + str(row['Commodity']) + ")"
    strikes_commodity_list.append(strike_commodity)
#add new columns to df
df["CStrike"] = strikes_commodity_list
df["Date"] = dates
#change all dashes to 0 for ease of adding total
df.loc[(df['Open Interest'] == '-'),'Open Interest']='0'
#remove commas in oi for ease of adding total
df['Open Interest'] = df['Open Interest'].replace({',':''}, regex=True)

#create a more ordered dictionary using the dataframe
for index, row in df.iterrows():
    date = str(row['Contract Month']) + "/" + str(row['Contract Year'])
    if row['Contract Type']=="P":
        try:
            finding = next(item for item in ndict[date] if item["strike"] == str(row['CStrike']))
            finding["put oi"] = int(row['Open Interest'])
        except:
            ndict[date].append({"strike": str(row['CStrike']), "put oi": int(row['Open Interest']), "call oi": 0})
    else:
        ndict[date].append({"strike": str(row['CStrike']), "call oi": int(row['Open Interest']), "put oi": 0})
#ndict would now be a dictionary with the keys as "month/year" and the values are lists of dictionaries
#eg. {"5/2020": [{"strike": "30.00 (FE)", "call oi": 0, "put oi": 0}, etc]}

# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

#create spreadsheet with SHEET_NAME as title
sh = client.create(SHEET_NAME)
#share new spreadsheet with all the emails in SHEET_EMAILS
for email in SHEET_EMAILS:
    sh.share(email, perm_type='user', role='writer')

# Find a spreadsheet by name and open the first sheet found
sheet = client.open(SHEET_NAME).sheet1

sheet.insert_row(unique_dates, 1) #inserts dates in first row
sheet.insert_row(second_row, 2) #inserts strike, call, put, total in second row

def change_col_char(start_char):
    '''Increments characters to specify column of cell range.

    Increments string start_char; from A-Z to AA-AZ to BA-BZ etc.
    Works for max up to ZZ.
    '''
    if len(start_char)==2: #if alr more than one character (starting character is A)
        first_char=start_char[0]
        start_char = start_char[1]
        if chr(ord(start_char) + 1)>"Z": #if at AZ go to BA
            numch = (ord(start_char) + 1)-26
            start_char= chr(ord(first_char) + 1)+chr(numch)
        else: #else add normally (two char with starting character A)
            start_char = chr(ord(start_char) + 1)
            start_char= first_char+start_char
    elif chr(ord(start_char) + 1)>"Z": #if it is at Z then go to AA
        numch = (ord(start_char) + 1)-26
        start_char= "A"+chr(numch)
    else: #else add normally (one char)
        start_char = chr(ord(start_char) + 1)
    return start_char


start_char = "A"
total_row = [] #for total row
for key, value in ndict.items():
    sleep(1.6) #to avoid error of exceeding 100 requests within 100s
    for elem in ["strike", "call oi", "put oi", "total"]:
        #append row number to column letters
        start = start_char+str(3)
        end = start_char+str(len(value)+3)
        #range of cells to insert data
        cell_list = sheet.range('%s:%s' % (start, end))
        iter1 = 0
        sum=0

        print(key+" "+elem) #to see progress in terminal
        #iterate through cell_list and ndict value to assign cell values
        for cell in cell_list:
            if iter1 in range(0, len(value)):
                if elem == "total":
                    cell.value = value[iter1]['put oi']+value[iter1]['call oi']
                    sum += value[iter1]['put oi']+value[iter1]['call oi']
                elif elem == "strike":
                    cell.value = value[iter1][elem]
                    sum = "Total"
                else:
                    cell.value = value[iter1][elem]
                    sum += value[iter1][elem]
                iter1 = iter1 + 1
        total_row.append(sum) #append sum of total call oi/put oi/total to total_row
        sheet.update_cells(cell_list)
        start_char = change_col_char(start_char)

sheet.insert_row(total_row, 95) #insert row for total oi in each column into row 95

ratio_row = [] #for put/call ratio row
put_i = 2
call_i = 1
while put_i in range(0, len(total_row)):
    ratio_row.extend(["Put/Call Ratio", ""])
    if (total_row[put_i]!=0 and total_row[call_i]!=0):
        ratio = round(total_row[put_i]/total_row[call_i],3)
    else:
        ratio = "-"
    ratio_row.extend([ratio, ""])
    put_i += 4
    call_i += 4
sheet.insert_row(ratio_row, 96) #insert row for put/call ratios for each month/year into row 96


format_sheet(SHEET_NAME, 0, SHEET_DATE, client) #to format the sheet