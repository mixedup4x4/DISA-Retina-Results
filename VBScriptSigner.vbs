' Change the following fields to suit your needs
' Const ftotest = "VBScriptToSign.vbs"
' oScrSig.SignFile sigfile, "SignedByWho"  
Option Explicit
Dim oScrSig, oFso, sigfile, sigstatus, scriptpath, showGUI
Const ftotest = "VBScriptToSign.vbs"

showGUI = True
set oScrSig = WScript.CreateObject("Scripting.Signer")
set oFso = WScript.CreateObject("Scripting.FileSystemObject")
scriptpath = oFso.GetParentFolderName(WSCript.ScriptFullName) & "\"

sigfile = scriptpath & ftotest

oScrSig.SignFile sigfile, "SignedByWho"    ' try to sign the file.
sigstatus = oScrSig.VerifyFile(sigfile, showGUI)    ' verify the signature.
 If sigstatus then
  WScript.Echo "Signature verified for " & ftotest
 Else
  WScript.Echo "Signature **FAILED** verification for " & ftotest
 End If
