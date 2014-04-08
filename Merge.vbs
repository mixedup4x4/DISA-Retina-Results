'******************************************************************
'*****Merge Multiple XML files together from several Scanners******
'*****Pieced together by some guy to automate the copy paste*******
'*****things that we do when it comes to merging XML files for*****
'*****uploading into VMS so there is a single bucket of SCCVI******
'******************************************************************
'*********************Created by Jason Chapell*********************
'******************************************************************

Option Explicit

'Defining Variables

Dim bln_Skip_IP
Dim bln_xml_Heading
Dim detCutoffTime			
Dim objFSO				
Dim colFiles				
Dim objFile				
Dim objFileInput
Dim objFolder				
Dim File_All_Retina			
Dim strCurAppPath
Dim strAll_Retina_Output			
Dim strFilename				
Dim str_I0_Folder			
Dim strOutputFldr
Dim strLine
Dim strNewAsset
Dim strPattern1				
Dim strPattern2				
Dim strDestinationPath

'Wait until msgBox displays showing complete so the file does not become corrupt
'Tell the user

msgBox("Please wait until the Message showing complete Otherwise, the file will become corrupt")

'Setting the Reading and Writing of the new XML

Const ForReading = 1		
Const ForWriting = 2		
Const ForAppending = 8		
strPattern1			= "xml"
strPattern2			= "xml"
strCurAppPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") 
str_I0_Folder = strCurAppPath
strOutputFldr = str_I0_Folder
strAll_Retina_Output	= strOutputFldr & "\" & "~Combined_Export.xml"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(str_I0_Folder)
Set colFiles = objFolder.Files
Set File_All_Retina 	= objFso.OpenTextFile(strAll_Retina_Output, ForWriting, true,0)
strLine = "<?xml version=" & chr(34) & "1.0" & chr(34) & " encoding=" & chr(34) & "ISO-8859-1"  & _
	chr(34) & " standalone=" & chr(34) & "yes" & chr(34) & " ?>"
File_All_Retina.WriteLine strLine
File_All_Retina.WriteLine "<IMPORT_FILE>"
File_All_Retina.Close

'Copying the individual xml files into the new combined XML

For Each objFile in colFiles
      If Left(objFile.name,1) <> "~" And _
	(Right(objFile.name, 3) = strPattern1 or Right(objFile.name, 3) = strPattern2) Then
		Set objFileInput = objFso.OpenTextFile(objFile.Path, ForReading)
		Set File_All_Retina = objFso.OpenTextFile(strAll_Retina_Output, ForAppending, False, 0)
		bln_Skip_IP   = False
		bln_xml_Heading = True
		Do While objFileInput.AtEndOfStream <> True
	   		strLine = objFileInput.ReadLine
			If Left(strLine,5) = "<?xml" Or Left(strLine,13) = "<IMPORT_FILE>" Or _
				Left(strLine, 14) = "</IMPORT_FILE>" Then
			Else
				File_All_Retina.WriteLine strLine
			End If       
		Loop
		File_All_Retina.Close
		objFileInput.Close
      Else
      End If
Next
Set File_All_Retina = objFso.OpenTextFile(strAll_Retina_Output, ForAppending, False, 0)
File_All_Retina.WriteLine "</IMPORT_FILE>"
File_All_Retina.Close

'Main process is completed
'The new XML has been created

MsgBox ("The Merging of XML Files is Complete")
MsgBox ("Don't Forget to ZIP before Upload")

 'Released under GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007
