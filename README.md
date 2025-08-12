VBA code is using trancript of SAP Script for T-code FBL5H from Client Accounting module.
The Workbook is having the client numbers in column A, folder path in cell D1 where the report will be saved and the file name in cell D2.
ctxtP_LAYOUT, ctxtDY_PATH and ctxtDY_FILENAME are modified with VBA variable input based on the Sheet data.
The extraction will be saved with cell D2 input & date of the extraction.

The SAP Script for the T-code can be optained by using Script record from SAP transaction Options and then insertd in the VBA code structure.
The VBA can be triggered directly with a VBS file command without oppening the Excel file where the VBA is storred. You can use the VBS:



'Input Excel File's Full Path
  ExcelFilePath = "C:\Users\ABCDE\Desktop\SAP_extraction.xlsm"

'Input Module/Macro name within the Excel File
  MacroPath = "Module1.Tcode_extraction"

'Create an instance of Excel
  Set ExcelApp = CreateObject("Excel.Application")

'Do you want this Excel instance to be visible?
  ExcelApp.Visible = False  'or "True"

'Prevent any App Launch Alerts (ie Update External Links)
  ExcelApp.DisplayAlerts = False

'Open Excel File
  Set wb = ExcelApp.Workbooks.Open(ExcelFilePath)

'Execute Macro Code
  ExcelApp.Run MacroPath

'Save Excel File (if applicable)
  wb.Save

'Reset Display Alerts Before Closing
  ExcelApp.DisplayAlerts = True

'Close Excel File
  wb.Close

'End instance of Excel
  ExcelApp.Quit

'Leaves an onscreen message!
  MsgBox "File generated at " & TimeValue(Now), vbInformation
