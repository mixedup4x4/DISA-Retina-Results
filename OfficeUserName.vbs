' This script will add the Users' correct Username and UserInitials to the MS Office
' registry key to prevent the first time prompt when a user runs up a program in the
' Office suite.
 
' Note that all Office versions before 2007 (12.0) used a Binary value, hence the
' reason for needing to use the ConvertStringToBinary() function.
 
' It gets the username and initials in one of two ways:
' 1) It first tries Active Directory to get the given and surname properties.
' 2) If that fails, it derives the initials from the logged on username based on
'    common naming standards.
'    For Example:
'             If the username is Jeremy.Saunders, the initials will be JS
'             If the username is jsaunders, the initals will also be JS
'    This is easy to change/add/modify should you be using a different naming standard
'    that follows a different pattern.
 
Option Explicit
 
Dim arrBinaryValue(), strUsername, strUserInitials, strTemp, intNumberOfChars, objWSHNetwork
Dim objShell, strComputer, objReg, strKeyRoot, strKeyPath, arrVersions, Version, return
Dim strUsernameInBinary, strUserInitialsInBinary, objSysInfo, strUserDN, objUserProperties
Dim blnDebug
 
Const HKEY_CURRENT_USER = &H80000001
 
' ********************** Set these variables *****************************
 
' Set this to True to help debug issues.
blnDebug = False
 
' Add the Office application versions you are using to the arrVersions array.
arrVersions = Array("10.0","11.0","12.0","14.0")
' Note that...
' - Office 2000 = 9.0
' - Office XP/2002 = 10.0
' - Office 2003 = 11.0
' - Office 2007 = 12.0
' - Office 2010 = 14.0
 
' ************************************************************************
 
strComputer = "."
strUsername = ""
strUserInitials = ""
 
Set objShell = WScript.CreateObject("WScript.Shell")
Set objWSHNetwork = WScript.CreateObject("WScript.Network")
Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
  strComputer & "\root\default:StdRegProv")
 
On Error Resume Next
' Get the user properties from Active Directory.
Set objSysInfo = CreateObject("ADSystemInfo")
If Err.Number = 0 Then
  strUserDN = objSysInfo.UserName
  Set objUserProperties = GetObject("LDAP://" & strUserDN)
  If Err.Number = 0 Then
    strUsername = objUserProperties.givenName & " " & objUserProperties.SN
    strUserInitials = Left(objUserProperties.givenName, 1) & Left(objUserProperties.SN, 1)
  Else
    If blnDebug Then
      wscript.echo "Cannot Retrieve User Properties from Active Directory."
    End If
  End If
Else
  If blnDebug Then
    wscript.echo "Cannot Connect to Active Directory."
  End If
End If
On Error Goto 0
Err.Clear
 
If strUsername="" Then
  strUsername = objWSHNetwork.UserName
  If instr(strUsername, ".") > 0 Then
    strTemp = Split(strUsername, ".")
    strUserInitials = ucase(Left(strTemp(0), 1)) & ucase(Left(strTemp(1), 1))
  Else
    strUserInitials = ucase(Left(strUsername, 2))
  End If
End If
 
If blnDebug Then
  wscript.echo "The username is: " & strUsername & vbcrlf & _
               "The initials are: " & strUserInitials
End If
 
If IsArray(arrVersions) Then
  For Each Version in arrVersions
    strKeyRoot = "HKCU\"
    strKeyPath = "Software\Microsoft\Office\"
 
    If Version = "9.0" OR Version = "10.0" OR Version = "11.0" Then
 
      strKeyPath = "Software\Microsoft\Office\" & Version
 
      If RegKeyExists(strKeyRoot & strKeyPath) Then
 
        If NOT RegKeyExists(strKeyRoot & strKeyPath & "\Common") Then
          return = objReg.CreateKey (HKEY_CURRENT_USER, strKeyPath & "\Common")
        End If
        If NOT RegKeyExists(strKeyRoot & strKeyPath & "\Common\UserInfo") Then
          return = objReg.CreateKey (HKEY_CURRENT_USER, strKeyPath & "\Common\UserInfo")
        End If
 
        strKeyPath = strKeyPath & "\Common\UserInfo"
 
        strUsernameInBinary = ConvertStringToBinary(strUsername)
        objReg.SetBinaryValue HKEY_CURRENT_USER, strKeyPath, "UserName", strUsernameInBinary
 
        strUserInitialsInBinary = ConvertStringToBinary(strUserInitials)
        objReg.SetBinaryValue HKEY_CURRENT_USER, strKeyPath, "UserInitials", strUserInitialsInBinary
 
      End If
 
    Else
 
      If RegKeyExists(strKeyRoot & strKeyPath & Version) Then
 
        If NOT RegKeyExists(strKeyRoot & strKeyPath & "\Common") Then
          return = objReg.CreateKey (HKEY_CURRENT_USER, strKeyPath & "\Common")
        End If
        If NOT RegKeyExists(strKeyRoot & strKeyPath & "\Common\UserInfo") Then
          return = objReg.CreateKey (HKEY_CURRENT_USER, strKeyPath & "\Common\UserInfo")
        End If
 
        strKeyPath = strKeyPath & "\Common\UserInfo"
 
        objReg.SetStringValue HKEY_CURRENT_USER, strKeyPath, "UserName", strUsername
        objReg.SetStringValue HKEY_CURRENT_USER, strKeyPath, "UserInitials", strUserInitials
 
      End If
    End If
  Next
End If
 
Set objWSHNetwork = Nothing
Set objShell = Nothing
Set objReg = Nothing
Set objSysInfo = Nothing
 
wscript.quit(0)
 
Function ConvertStringToBinary(strString)
  ReDim arrBinaryValue(len(strString) * 2 + 1)
  For intNumberOfChars = 0 To Len(strString) - 1
    If intNumberOfChars = 0 Then
      arrBinaryValue(0) = Asc(Mid(strString, intNumberOfChars + 1, 1))
      arrBinaryValue(1) = 0
    Else
      arrBinaryValue(intNumberOfChars * 2) = Asc(Mid(strString, intNumberOfChars + 1, 1))
      arrBinaryValue(intNumberOfChars * 2 + 1) = 0
    End If
  Next
  arrBinaryValue(Len(strString) * 2) = 0
  arrBinaryValue(Len(strString) * 2 + 1) = 0
  ConvertStringToBinary = arrBinaryValue
End Function
 
Function RegKeyExists(ByVal sRegKey)
' Returns True or False based on the existence of a registry key.
  Dim sDescription, oShell
  Set oShell = CreateObject("WScript.Shell")
  RegKeyExists = True
  sRegKey = Trim (sRegKey)
  If Not Right(sRegKey, 1) = "\" Then
    sRegKey = sRegKey & "\"
  End If
  On Error Resume Next
  oShell.RegRead "HKEYNotAKey\"
  sDescription = Replace(Err.Description, "HKEYNotAKey\", "")
  Err.Clear
  oShell.RegRead sRegKey
  RegKeyExists = sDescription <> Replace(Err.Description, sRegKey, "")
  On Error Goto 0
  Set oShell = Nothing
End Function
