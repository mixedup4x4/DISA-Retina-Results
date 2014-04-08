Attribute VB_Name = "Module5"
Sub Protect_Selected_Sheets()
Attribute Protect_Selected_Sheets.VB_ProcData.VB_Invoke_Func = "l\n14"
Dim wb As Workbook
Dim ws As Worksheet
Dim blnIsProtected As Boolean
Set wb = ActiveWorkbook

For Each ws In wb.Worksheets
ws.Protect "Password123"
Next ws

Set wb = Nothing
Set ws = Nothing
End Sub
