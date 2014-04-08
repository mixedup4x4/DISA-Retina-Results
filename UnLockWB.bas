Attribute VB_Name = "Module6"

Sub Unprotect_selected_sheets()
Attribute Unprotect_selected_sheets.VB_ProcData.VB_Invoke_Func = "q\n14"
Dim wb As Workbook
Dim ws As Worksheet
Dim blnIsProtected As Boolean
Set wb = ActiveWorkbook

For Each ws In wb.Worksheets
ws.Unprotect "Password123"
Next ws

Set wb = Nothing
Set ws = Nothing
End Sub
