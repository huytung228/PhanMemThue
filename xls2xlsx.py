import os
import os.path
import win32com.client as win32
 
# ## Root directory 
# rootdir = r'C:\Users\zz\Desktop\PhanMemThue\data\5_3b_21XX_temp'
#  # three parameters: parent directory; all folder names (without path); all file names
# for parent, dirnames, filenames in os.walk(rootdir):
#     for fn in filenames:
#         filedir = os.path.join(parent, fn)
#         print(filedir)
 
#         excel = win32.gencache.EnsureDispatch('Excel.Application')
#         wb = excel.Workbooks.Open(filedir)
#         # xlsx: FileFormat=51
#         # xls:  FileFormat=56,
#         wb.SaveAs(filedir.replace('XLS', 'xlsx'), FileFormat=51)
#         wb.Close()                                 
#         excel.Application.Quit()

def xls2xlsx(input_dir):
    for xls_file in os.listdir(input_dir):
        if '.xls' in xls_file or '.XLS' in xls_file:
            filedir = os.path.join(input_dir, xls_file)
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(filedir)
            # xlsx: FileFormat=51
            # xls:  FileFormat=56,
            wb.SaveAs(filedir.replace('XLS', 'xlsx'), FileFormat=51)
            wb.Close()                                 
            excel.Application.Quit()

if __name__ == '__main__':
    xls2xlsx(r'C:\Users\zz\Desktop\PhanMemThue\data\5_3b_21XX_temp')