import win32com.client

o = win32com.client.Dispatch("Excel.Application")



wb_path = r'C:\Users\supaw\PycharmProjects\ConvertToPython\October2020Summary.xlsx'

wb = o.Workbooks.Open(wb_path)



ws_index_list = [1] #say you want to print these sheets

path_to_pdf = r'Desktop\sample.pdf'



wb.WorkSheets(ws_index_list).Select()

wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)