#%% open sheet
import win32com.client as win32
import os
xl = win32.DispatchEx('Excel.Application')
xl.DisplayAlerts = False
xl.Visible = True
file_path = r"C:\Code\XLRDTest\testfile.xls"
file_out = r"C:\Code\XLRDTest\testfile_out.xls"
wb = xl.Workbooks.Open(Filename=os.path.abspath(file_path))
sheet = wb.Sheets(1)
#%% modify

sheet.Cells(1,1).value = 53
sheet.Cells(3,1).EntireRow.Delete(Shift=-4162)


#%%
print(sheet.Cells(1,1).value)
sheet.Range('A1').EntireRow.Delete(Shift=-4162)
wb.SaveAs(file_out)
wb.Close()
xl.Quit()
# %%
