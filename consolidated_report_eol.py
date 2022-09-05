""""
1. to open the excel sheet and be able to extract the data from the column
2. create a loop for the
to read the excel sheet
pick those value and feed in respective cycle
find the total number of li

"""
from email import header
import sys
import xlsxwriter
import time
import os
import xlrd 


# the location of the file 
loc = r"C:\Users\ajana\Desktop\python-project\OUTPUT_16_08_2022_eol_setup_asc_logs1_1.xlsx"
loc1 = r"C:\Users\ajana\Desktop\python-project\testing_scripts"
excel_sheet_list = os.listdir(loc1)
"""
11_08_2022_NHW_23.40

"""
date = []
build_hardware = []
build_no = []

for f in excel_sheet_list:
    date.append(f[0:10])
    build_hardware.append(f[11:14])
    build_no.append(f[15:20])



OUTPUT="OUTPUT_consolidated_report"+".xlsx"
workbook = xlsxwriter.Workbook(OUTPUT)

worksheet_version_history = workbook.add_worksheet("Version History")
worksheet_eol_summary = workbook.add_worksheet("EOL summary")

# setting up the version history for the excel files
border_format = workbook.add_format({
    'border': 2,
    'align':'left',
    'font_size': 15
})
worksheet_version_history.write('C2','Project Name',border_format)
worksheet_version_history.write('C3','Document Title',border_format)
worksheet_version_history.write('C4','Date',border_format)
worksheet_version_history.write('C5','Version',border_format)
worksheet_version_history.write('C6','Status',border_format)
worksheet_version_history.write('C7','Owner',border_format)
worksheet_version_history.write('C8','Approved',border_format)


worksheet_version_history.write('B12','Version',border_format)
worksheet_version_history.write('C12','Date',border_format)
worksheet_version_history.write('D12','Change from Previous',border_format)


cnt_version_row = 13

for i in range(len(excel_sheet_list)):
    worksheet_version_history.write(cnt_version_row,1,(i+1)/10,border_format)
    worksheet_version_history.write(cnt_version_row,2,date[i],border_format)
    title = "EOL test report on " + build_hardware[i] + build_no[i]
    worksheet_version_history.write(cnt_version_row,3,title,border_format)

# setting up the eol sumary for the excel files

header_format=workbook.add_format({'top':1,'right':1,'left':1,'bottom':1,'align':'center'})
header_format.set_bg_color("yellow")

cnt_summary_row = 3
for i in range(len(excel_sheet_list)):
    lr = 'B'+str(cnt_summary_row)
    hr = 'C'+str(cnt_summary_row)
    tr = lr+":"+hr
    worksheet_eol_summary.merge_range(tr,'PSA CARUSO Test Case Execution Summary',header_format)
    cnt_summary_row += 1
    worksheet_eol_summary.write(cnt_summary_row,1,'EOL Tool Version',header_format)
    worksheet_eol_summary.write(cnt_summary_row,2,'0.2',header_format)
    cnt_summary_row +=1
    worksheet_eol_summary.write(cnt_summary_row,1,'Brand',header_format)
    worksheet_eol_summary.write(cnt_summary_row,2,build_hardware[i],header_format)
    cnt_summary_row += 1
    worksheet_eol_summary.write(cnt_summary_row,1,'Build Version',header_format)
    worksheet_eol_summary.write(cnt_summary_row,2,build_no[i],header_format)
    cnt_summary_row +=1
    worksheet_eol_summary.write(cnt_summary_row,1,'Number of cycles',header_format)
    cnt_summary_row += 5











# for putting the timings for 


for files in excel_sheet_list:
    # new_loc = os.path.abspath(files)
    new_loc = os.path.join(loc1,files)
    worksheet = workbook.add_worksheet(files)
    wb = xlrd.open_workbook(new_loc) 
    sheet = wb.sheet_by_index(0) 
    rows = sheet.nrows 
    total_no_of_dids = rows - 3
    no_of_dids_in_one_cycle = 1



    for i in range(3,total_no_of_dids): 
        if(sheet.cell_value(i, 4)=="Fault Memory Read (by Mask)" and sheet.cell_value(i+1,4)=="Default session"):
            break
        else:
            no_of_dids_in_one_cycle += 1


    print("no of dids in one cycle ",no_of_dids_in_one_cycle)
    total_no_of_cycles = total_no_of_dids/no_of_dids_in_one_cycle
    print(total_no_of_cycles)
    no_of_cycles = int(total_no_of_cycles)
    did_cnt_idx = 3
    print(no_of_cycles)

    print(sheet.cell_value(did_cnt_idx,5))
    
    cnt = 0
    for dids in range(3,no_of_dids_in_one_cycle+3):
        did_values = sheet.cell_value(dids,4)
        worksheet.write(cnt,0,did_values)
        cnt+=1

    for c in range(2,no_of_cycles+2):
        for dids in range(1,no_of_dids_in_one_cycle+1):
            time_taken = sheet.cell_value(did_cnt_idx,5)
            cyc_header = "cycle" + str(c-1)
            worksheet.write(0,c,cyc_header)
            worksheet.write(dids,c,time_taken)
            if(did_cnt_idx != total_no_of_dids):
                did_cnt_idx += 1
            else:
                break



workbook.close()

