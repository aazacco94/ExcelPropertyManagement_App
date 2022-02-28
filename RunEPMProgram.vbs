Set objExcel = CreateObject("Excel.Application")
objExcel.Application.Run "'C:\Users\zacco\School\SER416\FinalProject\Program\EPM_Program.xlsm'!Main"
objExcel.DisplayAlerts = False
objExcel.Application.Quit
Set objExcel = Nothing