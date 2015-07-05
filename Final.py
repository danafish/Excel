from win32com import client
import os
import sys

def getScriptPath():
    return os.path.dirname(os.path.realpath(sys.argv[0]))


file_input=getScriptPath()
#raw_input("pleas copy the path of your main directory and paste here \n")
print file_input

for file in os.listdir(file_input):
    if file.endswith("Charges.xls")or file.endswith("Charges.xlsx"):
        print('this is the main file:'+file)
        print 'is this the right file? write Yes\\No \n'
        if raw_input("Yes\No: ")=='Yes':
            main_file=file
            break
if not main_file:
   print 'no filethat ends with Charges.xlsx or Carges.xls'

#file name
end_input='template.xlsx'

#file output path
file_output=os.makedirs(file_input+'\\''%sBills'%main_file[:8])
file_output_name=file_input+'\\'+'%sBills' %main_file[:8]
disp=client.Dispatch("Excel.Application")
wb_s=disp.Workbooks.Open(file_input+'\\'+main_file)
wb_d=disp.Workbooks.Open(file_input+'\\'+end_input)
#runs over the sheets
#s=source, d=destination, t=temporary, ws=work sheet, wb=work book
ws_s=wb_s.Sheets('Summary')
ws_d = wb_d.Worksheets("Sheet2")
#opens and saves excel for each row to sheet
row=3
print 'here we go:'
try:
    while ws_s.Range('A%d'%row).value:
        temp_file_name = ws_s.Range('O%d' %row).value
        print temp_file_name
        if temp_file_name:
            ws_t = wb_s.Worksheets(str(row-2))
            ws_t.Range('A1:J25').Copy()
            ws_d.Paste(ws_d.Range('A1'))
            wb_d.WorkSheets(["Sheet1", "Sheet2"]).select
            ws_d.SaveAs(file_output_name+'\\'+temp_file_name+'.xlsx')
    #saves as pdf only 2 relevant pages
            ws_d.ExportAsFixedFormat(0 , '%s\\%s.pdf' %(file_output_name,temp_file_name),0,True,True,1,2)
        row+=1
except:
    print 'sorry didnt work'
disp.Application.Quit()
print 'We are Done. YAY :)'