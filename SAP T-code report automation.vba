Sub Tcode_extraction()

Dim sht As Worksheet
Dim last_row As Long
Dim StartCell As Range
Set sht = Worksheets("Sheet1")
Set StartCell = Range("A1")

last_row = sht.Cells(sht.Rows.Count, StartCell.Column).End(xlUp).Row
If last_row > 1 Then

If Not IsObject(APP) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set APP = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = APP.Children(0)
End If
If Not IsObject(session) Then
   Set session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject APP, "on"
End If

session.findById("wnd[0]/tbar[0]/okcd").Text = "/nfbl5h"
session.findById("wnd[0]").sendVKey 0

Sheet1.Range("A2:A" & last_row).Copy

session.findById("wnd[0]/usr/btn%_S_CUST_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtP_LAYOUT").Text = "EXT"
session.findById("wnd[0]/usr/ctxtP_LAYOUT").SetFocus
session.findById("wnd[0]/usr/ctxtP_LAYOUT").caretPosition = 5
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGC_CONTAINER/shellcont/shell/shellcont[0]/shell").setCurrentCell -1, ""
session.findById("wnd[0]/usr/cntlGC_CONTAINER/shellcont/shell/shellcont[0]/shell").SelectAll
session.findById("wnd[0]/usr/cntlGC_CONTAINER/shellcont/shell/shellcont[0]/shell").pressToolbarButton "REPORT_CALL_LINE_ITEM"
session.findById("wnd[0]/usr/lbl[9,8]").SetFocus
session.findById("wnd[0]/usr/lbl[9,8]").caretPosition = 3
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[41]").press
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Sheet1.Range("D1").Value
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = Sheet1.Range("D2").Value & DateString & ".XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 7
session.findById("wnd[1]/tbar[0]/btn[0]").press
End If

End Sub
