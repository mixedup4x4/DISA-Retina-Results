Attribute VB_Name = "Module1"
Sub SplitSheets()
Attribute SplitSheets.VB_ProcData.VB_Invoke_Func = "S\n14"
Dim W As Worksheet
For Each W In Worksheets
W.SaveAs ActiveWorkbook.Path & "/" & W.Name
Next W
End Sub
