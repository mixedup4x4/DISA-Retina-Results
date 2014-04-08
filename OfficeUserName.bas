Attribute VB_Name = "Module3"
Sub OfficeUserName()
Attribute OfficeUserName.VB_ProcData.VB_Invoke_Func = "U\n14"
Dim RetVal
RetVal = Shell("PATH\TO\OfficeUserName.vbs", 0)
MsgBox ("UserName for Office Applications Modified") & vbOK
End Sub
