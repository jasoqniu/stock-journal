import openpyxl
from openpyxl import Workbook
from openpyxl import workbook
from openpyxl import load_workbook
import yfinance as yf
import datetime 
from datetime import datetime as dtm
import investpy
import pandas as pd
from datetime import timedelta
from openpyxl.styles import PatternFill
from openpyxl.styles import numbers
from openpyxl.styles import Font, Color

# set path
path = "C:\\Users\\Jason\\Downloads\\puefolder\\"

#set excel workbook
wb_AL = openpyxl.Workbook()
wb_MZ = openpyxl.Workbook()
wb_AL.remove(wb_AL["Sheet"])
wb_MZ.remove(wb_MZ["Sheet"])


#set date & time
current_datetime = dtm.now()
time_delta = timedelta(days=162)
past = current_datetime - time_delta
current_date = current_datetime.strftime("%Y-%m-%d")
date = past.strftime("%d-%m-%Y")

#set stocklist from investpy
#stock_list = investpy.get_stocks_list(country= "Indonesia")
#stock_list.sort()

#STOCKLIST = https://www.idx.co.id/id/data-pasar/data-saham/daftar-saham/
excel_file = path + 'stocklist.xlsx'
df = pd.read_excel(excel_file)
stock_list = df['Kode'].tolist()


#looping stock list
for s in stock_list:
    kode = s
    first_letter = s[0]
    print("Symbol is : ",s)
    data = yf.download(kode + ".JK", start=past,end=current_datetime)
    data = data.sort_index(ascending=False)
    
    #filter volume under 1 000 000 000
    avg_volume_100 = data['Volume'].mean()
    avg_close_100 = data['Close'].mean()

    if avg_volume_100 * avg_close_100 < 1000000000:
        print(kode,"removed, volume under 1 000 000 000")
        continue
    
    #filter into A-L and M-Z
    if 'A' <= first_letter <= 'L':
        current_workbook = wb_AL
    elif 'M' <= first_letter <= 'Z':
        current_workbook = wb_MZ
    
    
    #create new sheet
    new_sheet = current_workbook.create_sheet(kode)
    sheet = current_workbook[kode]
    
    #data length
    data_length = len(data)
    print("AMOUNT OF DAYS = ",data_length)
    
    #set label for row=1
    sheet.cell(row=1, column=1).value = "Tanggal"
    sheet.cell(row=1, column=2).value = "Open"
    sheet.cell(row=1, column=3).value = "High"
    sheet.cell(row=1, column=4).value = "Low"
    sheet.cell(row=1, column=5).value = "Close"
    sheet.cell(row=1, column=6).value = "Volume"
    sheet.cell(row=1, column=8).value = "CH"
    sheet.cell(row=1, column=9).value = "CL"
    sheet.cell(row=1, column=10).value = "CC"
    sheet.cell(row=1, column=11).value = "Avg Harian"
    sheet.cell(row=1, column=12).value = "MA5"
    sheet.cell(row=1, column=13).value = "op=low"
    sheet.cell(row=1, column=14).value = "op=high"
    sheet.cell(row=1, column=15).value = "Prank"
    sheet.cell(row=1, column=16).value = "JJS OpLo"
    sheet.cell(row=1, column=17).value = "WR JJSOL"
    sheet.cell(row=1, column=18).value = "Prank%"
    sheet.cell(row=1, column=19).value = "OpLo9"
    sheet.cell(row=1, column=20).value = "WR OpLo9"
    
    
    #set width for row=1
    sheet.column_dimensions['A'].width = 15
    sheet.column_dimensions['B'].width = 10
    sheet.column_dimensions['C'].width = 10
    sheet.column_dimensions['D'].width = 10
    sheet.column_dimensions['E'].width = 10
    sheet.column_dimensions['F'].width = 13

    #color row=1
    for c in range(1,21):
        sheet.cell(row=1,column=c).fill = PatternFill(start_color="94c685",
                                        end_color="94c685",
                                        fill_type="solid")
        sheet.cell(row=1,column=c).font = Font(color="FFFFFF", bold=True, italic=False)
    
    #OHLCV
    rowvalue = 2
    for index,row in data.iterrows():
        open_value = row['Open']
        high_value = row['High']
        low_value = row['Low']
        close_value = row['Close']
        volume_value = row['Volume']
        tanggal_value = index.date()
        '''
        open_position = sheet.cell(row=rowvalue, column=2)
        high_position = sheet.cell(row=rowvalue, column=3) 
        low_position = sheet.cell(row=rowvalue, column=4) 
        close_position = sheet.cell(row=rowvalue, column=5)
        volume_position = sheet.cell(row=rowvalue, column=6)
        '''
        sheet.cell(row=rowvalue, column=1).value = tanggal_value
        #sheet.cell(row=rowvalue, column=1).number_format = numbers.FORMAT_DATE_XLSX14
        sheet.cell(row=rowvalue, column=2).value = open_value
        sheet.cell(row=rowvalue, column=3).value = high_value
        sheet.cell(row=rowvalue, column=4).value = low_value
        sheet.cell(row=rowvalue, column=5).value = close_value
        sheet.cell(row=rowvalue, column=6).value = volume_value
        rowvalue += 1
    
    #Color OHLCV
    for rowvalue in range(2, (data_length+2)):
        if rowvalue % 2 == 0:
            for colvalue in range(1,7):
                cell = sheet.cell(row=rowvalue,column=colvalue)
                cell.fill = PatternFill(start_color="ffffff",
                                        end_color="ffffff",
                                        fill_type="solid")
        else:
            for colvalue in range(1,7):
                cell = sheet.cell(row=rowvalue,column=colvalue)
                cell.fill = PatternFill(start_color="a2c997",
                                        end_color="a2c997",
                                        fill_type="solid")        
    
    #CH
    rowvalue = 2
    for index,row in data.iterrows():
        high_today = row['High']
        close_yesterday = sheet.cell(row=rowvalue + 1, column=5).value
        if close_yesterday:
            ch_value = (high_today - close_yesterday) / close_yesterday
            sheet.cell(row=rowvalue, column=8).value = ch_value
            #data.at[index, 'CH'] = ch_value
        sheet.cell(row=rowvalue, column=8).number_format = numbers.FORMAT_PERCENTAGE_00
        rowvalue += 1
    
    #CL
    rowvalue = 2
    for index, row in data.iterrows():
        low_today = row['Low']
        close_yesterday = sheet.cell(row=rowvalue + 1, column=5).value
        if close_yesterday:
            cl_value = (low_today-close_yesterday) / close_yesterday
            sheet.cell(row=rowvalue, column=9).value = cl_value
        sheet.cell(row=rowvalue, column=9).number_format = numbers.FORMAT_PERCENTAGE_00
        rowvalue += 1

    #CC
    rowvalue = 2
    for index, row in data.iterrows():
        close_yesterday = sheet.cell(row=rowvalue + 1, column=5).value
        close_today = sheet.cell(row=rowvalue, column=5).value
        if close_yesterday:
            cc_value = (close_today - close_yesterday) / close_yesterday
            sheet.cell(row=rowvalue, column=10).value = cc_value
        sheet.cell(row=rowvalue, column=10).number_format = numbers.FORMAT_PERCENTAGE_00
        rowvalue += 1
    
    #Color CH CL CC
    rowvalue = 2
    for r in range(data_length):
        ch_value = sheet.cell(row=rowvalue, column=8).value 
        cl_value = sheet.cell(row=rowvalue, column=9).value 
        cc_value = sheet.cell(row=rowvalue, column=10).value 

        if ch_value is not None and ch_value >= 0.02:
            sheet.cell(row=rowvalue,column=8).fill = PatternFill(start_color="6fe26f",
                                                                 end_color="6fe26f",
                                                                   fill_type="solid")
        if cl_value is not None and cl_value <= -0.03:
            sheet.cell(row=rowvalue,column=9).fill = PatternFill(start_color="e2746f",
                                                                 end_color="e2746f",
                                                                   fill_type="solid")
            sheet.cell(row=rowvalue, column=9).font = Font(color="FFFFFF", bold=True, italic=False)
            
        if cc_value is not None and cc_value <= -0.03:
             sheet.cell(row=rowvalue,column=10).fill = PatternFill(start_color="e2746f",
                                                                 end_color="e2746f",
                                                                   fill_type="solid")
             sheet.cell(row=rowvalue, column=10).font = Font(color="FFFFFF", bold=True, italic=False)
        rowvalue += 1
        
    
    #Avg Harian
    rowvalue = 2
    for index,row in data.iterrows():
        open_value = row['Open']
        high_value = row['High']
        low_value = row['Low']
        close_value = row['Close']
        avg_harian = (open_value + high_value + low_value + close_value) / 4
        sheet.cell(row=rowvalue,column=11).value = avg_harian
        rowvalue += 1
    
    #MA5
    for r in range(2,(data_length+2)):
        if r <= (data_length-4):
            sum_ma5 = 0
            for i in range(5):
                sum_ma5 += sheet.cell(row=r+i,column=11).value
            ma5_value = sum_ma5 / 5
            sheet.cell(row=r,column=12).value = ma5_value
    
    #op=lo op=high prank
    prank_true = 0
    prank_false = 0
    for r in range(2,(data_length+2)):
        open_value = sheet.cell(row=r,column=2).value
        high_vaue = sheet.cell(row=r,column=3).value
        low_value = sheet.cell(row=r,column=4).value
        ch_value = sheet.cell(row=r, column=8).value 

        
        #op==lo
        if sheet.cell(row=r,column=2).value == sheet.cell(row=r,column=4).value:
            sheet.cell(row=r,column=13).value = "YES"
        else:
            sheet.cell(row=r,column=13).value = "---"
            
        #op==high
        if  sheet.cell(row=r,column=2).value == sheet.cell(row=r,column=3).value:
            sheet.cell(row=r,column=14).value = "YES"
            
            #prank
            if sheet.cell(row=r, column=8).value is not None and sheet.cell(row=r, column=8).value >= 0.02:
                sheet.cell(row=r,column=15).value = "PRANK"
                prank_true += 1
                #print("PRANK = TRUE")
            else:
                sheet.cell(row=r,column=15).value = "---"
                prank_false +=1
            
        else:
            sheet.cell(row=r,column=14).value = "---"
        
    #prank%
    prank_total = prank_true + prank_false
    if prank_true == 0:
        prank_percent = 0
    else:
        prank_percent = prank_true / prank_total
    #print("PRANK% = ",prank_percent)
    sheet.cell(row=2, column=18).value  = prank_percent
    sheet.cell(row=2, column=18).number_format = numbers.FORMAT_PERCENTAGE_00       
        
    #JJSOplo 
    jjsoplo_wins = 0
    jjsoplo_count = 0
    for r in range(2,(data_length+2)):
        open_today = sheet.cell(row=r,column=2).value
        low_today = sheet.cell(row=r,column=4).value  
        ch_tomorrow = sheet.cell(row=r-1, column=8).value 
        
        if open_today == low_today and r != 2:
            jjsoplo_count +=1
            if ch_tomorrow is not None and ch_tomorrow >= 0.02:
                sheet.cell(row=r,column=16).value = "WIN"
                jjsoplo_wins += 1
            else:
                sheet.cell(row=r,column=16).value = "LOOSE"
        else:
            sheet.cell(row=r,column=16).value = "---"
    #JJSOplo WR%        
    if jjsoplo_wins == 0:
        jjsoplo_wr = 0
    else:
        jjsoplo_wr = jjsoplo_wins / jjsoplo_count
    sheet.cell(row=2,column=17).value = jjsoplo_wr
    sheet.cell(row=2,column=17).number_format = numbers.FORMAT_PERCENTAGE_00

    #OpLo9
    oplo9_win = 0
    oplo9_count = 0
    for r in range(2,(data_length+2)):
        open_value = sheet.cell(row=r,column=2).value
        high_value = sheet.cell(row=r,column=3).value
        low_value = sheet.cell(row=r,column=4).value
        
        if open_value == low_value:
            oplo9_count +=1
            if (high_value / open_value) > 1.03:
                sheet.cell(row=r,column=19).value = "WIN"
                oplo9_win +=1
            else:
                sheet.cell(row=r,column=19).value = "LOOSE"
        else:
            sheet.cell(row=r,column=19).value = "---"
    
    #WR OpLo9
    if oplo9_win == 0:
        oplo9_wr = 0
    else:
        oplo9_wr = oplo9_win/oplo9_count
    sheet.cell(row=2,column=20).value = oplo9_wr
    sheet.cell(row=2,column=20).number_format = numbers.FORMAT_PERCENTAGE_00
    
#save to A-L & M-Z workbook
wb_AL.save(path + "new_AL.xlsx")
wb_MZ.save(path + "new_MZ.xlsx")
    
    
#----------FILTERING----------#

#///FILTER JJSOL///
def filter_wrjjsol(wb, arr):
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        
        vol_sum = 0
        for r in range(2, 102):
            if sheet.cell(row=r, column=6).value is not None:
                vol_sum += sheet.cell(row=r, column=6).value
        vol_avg = vol_sum / 100
        
        if vol_avg < 50000:
            continue
        
        wrjjsol = sheet.cell(row=2, column=17).value

        if wrjjsol > 0.6:
            arr.append((sheet_name, wrjjsol,))
# arr for filtered stocks
arr_filter_wrrjjsol = []
# call function
filter_wrjjsol(wb_AL, arr_filter_wrrjjsol)
filter_wrjjsol(wb_MZ, arr_filter_wrrjjsol)

# sort stock
arr_filter_wrrjjsol.sort(key=lambda x: x[1], reverse=True)

# create new wb
wb_tugas = Workbook()
wb_tugas.remove(wb_tugas["Sheet"])
new_sheet = wb_tugas.active
new_sheet = wb_tugas.create_sheet("WRJJSOL")
new_sheet.append(["Kode", "WR% JJSOL"])

rowvalue = 2
for kode, wr in arr_filter_wrrjjsol:
    new_sheet.append([kode, wr])
    new_sheet.cell(row=rowvalue, column=2).number_format = numbers.FORMAT_PERCENTAGE_00
    rowvalue+= 1
# save 
#wb_tugas.save(path + "filtered_stocks_combined.xlsx")


#///FILTER OPLO9///#
def fiter_wroplo9(wb, arr):
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        
        vol_sum = 0
        for r in range(2, 102):
            if sheet.cell(row=r, column=6).value is not None:
                vol_sum += sheet.cell(row=r, column=6).value
        vol_avg = vol_sum / 100
        
        if vol_avg < 50000:
            continue
        
        wroplo9 = sheet.cell(row=2, column=20).value
        if wroplo9 > 0.6:
            arr.append((sheet_name, wroplo9,))
# arr for filtered stocks
arr_filter_wroplo9 = []
# call function
fiter_wroplo9(wb_AL, arr_filter_wroplo9)
fiter_wroplo9(wb_MZ, arr_filter_wroplo9)

# sort stock
arr_filter_wroplo9.sort(key=lambda x: x[1], reverse=True)

# create new wb
new_sheet = wb_tugas.create_sheet("WROPLO9")
new_sheet.append(["Kode", "WR% OpLo9"])

rowvalue = 2
for kode, wr in arr_filter_wroplo9:
    new_sheet.append([kode, wr])
    new_sheet.cell(row=rowvalue, column=2).number_format = numbers.FORMAT_PERCENTAGE_00
    rowvalue+= 1
# save 
wb_tugas.save(path + "filtered_stocks_combined.xlsx")


#-----------------------------#
