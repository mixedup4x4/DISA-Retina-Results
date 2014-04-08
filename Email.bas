Attribute VB_Name = "Module2"
Sub EmailProgress()
Attribute EmailProgress.VB_ProcData.VB_Invoke_Func = "E\n14"
Dim CDO_Mail As Object
Dim CDO_Config As Object
Dim SMTP_Config As Variant
Dim strSubject As String
Dim strFrom As String
Dim strTo As String
Dim strCc As String
Dim strBcc As String
Dim strBody As String

strSubject = "Documents have been updated"""
strFrom = "me@myself.com"
strTo = "you@yourself.com"
strCc = ""
strBcc = ""
strBody = ""

Set CDO_Mail = CreateObject("CDO.Message")
On Error GoTo Error_Handling

Set CDO_Config = CreateObject("CDO.Configuration")
CDO_Config.Load -1

Set SMTP_Config = CDO_Config.Fields

With SMTP_Config
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.something.com"
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    .Update
End With

With CDO_Mail
    Set .Configuration = CDO_Config
End With

CDO_Mail.Subject = strSubject
CDO_Mail.From = strFrom
CDO_Mail.To = strTo
CDO_Mail.TextBody = strBody
CDO_Mail.CC = strCc
CDO_Mail.BCC = strBcc
CDO_Mail.Send

Error_Handling:
If Err.Description <> "" Then MsgBox Err.Description
End Sub
