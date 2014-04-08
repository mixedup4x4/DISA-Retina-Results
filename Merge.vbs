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
'' SIG '' Begin signature block
'' SIG '' MIIOdAYJKoZIhvcNAQcCoIIOZTCCDmECAQExDjAMBggq
'' SIG '' hkiG9w0CBQUAMGYGCisGAQQBgjcCAQSgWDBWMDIGCisG
'' SIG '' AQQBgjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIB
'' SIG '' AAIBAAIBAAIBAAIBADAgMAwGCCqGSIb3DQIFBQAEENpR
'' SIG '' gdHOVgfXarrPkWK4Jl+gggwcMIIF4DCCBMigAwIBAgIR
'' SIG '' AKre8vwBbNKdvVNlITHpepgwDQYJKoZIhvcNAQELBQAw
'' SIG '' MzELMAkGA1UEBhMCVVMxEDAOBgNVBAoTB09SQyBQS0kx
'' SIG '' EjAQBgNVBAMTCU9SQyBTU1AgMzAeFw0xMzEwMTgxNjAw
'' SIG '' NTlaFw0xNjEwMTcxNjAwNTlaMIGAMQswCQYDVQQGEwJV
'' SIG '' UzEYMBYGA1UEChMPVS5TLiBHb3Zlcm5tZW50MQ4wDAYD
'' SIG '' VQQLEwVVU0VQQTEOMAwGA1UECxMFU3RhZmYxIjAgBgNV
'' SIG '' BAMTGUpBU09OIENIQVBFTEwgKGFmZmlsaWF0ZSkxEzAR
'' SIG '' BgNVBC4TCjAwMDAwNjQyMDIwggEiMA0GCSqGSIb3DQEB
'' SIG '' AQUAA4IBDwAwggEKAoIBAQDHwIJo7XONS4ivtkVONNyD
'' SIG '' GdcaSS/QWFEtNuTcVBuaoeiwYiLkBQZCpkHuZmViJkcu
'' SIG '' USEUeUQu4kdXoVKKVPCjOKgKkXkyKrFNTJs/J7qjzIxZ
'' SIG '' HA/yVBognzgR0gBOJYRcI1b2GCqG34oSVxqk1swGKxd2
'' SIG '' 6cXRrcZy9aqYhhTcyDYE+t23yFwuHXOfUkD137IiXFKz
'' SIG '' GbNaYkVr2hKekf4SSUZEWTDTfbcNJfMB7YPxMtUByG0A
'' SIG '' przyb9LF+9EVBvMXlY3WOw+1aqO2ANTG1A7itHAVd8hy
'' SIG '' SoSHOGHElW0g78DnSImXmwE3upAULrTiX90jB/gxqrKq
'' SIG '' tYgIKeEslJlXAgMBAAGjggKfMIICmzAfBgNVHSMEGDAW
'' SIG '' gBRqiG1S/LFEnjCuMxhNwDmdlmsksDAlBgNVHSUEHjAc
'' SIG '' BgRVHSUABgorBgEEAYI3FAICBggrBgEFBQcDAjAOBgNV
'' SIG '' HQ8BAf8EBAMCAIAwgfEGCCsGAQUFBwEBBIHkMIHhMCQG
'' SIG '' CCsGAQUFBzABhhhodHRwOi8vc3NwMy5ldmEub3JjLmNv
'' SIG '' bS8wOQYIKwYBBQUHMAKGLWh0dHA6Ly9jcmwtc2VydmVy
'' SIG '' Lm9yYy5jb20vY2FDZXJ0cy9PUkNTU1AzLnA3YzB+Bggr
'' SIG '' BgEFBQcwAoZybGRhcDovL29yYy1kcy5vcmMuY29tL2Nu
'' SIG '' JTNkT1JDJTIwU1NQJTIwMyUyY28lM2RPUkMlMjBQS0kl
'' SIG '' MmNjJTNkVVM/Y0FDZXJ0aWZpY2F0ZTtiaW5hcnksY3Jv
'' SIG '' c3NDZXJ0aWZpY2F0ZVBhaXI7YmluYXJ5MBcGA1UdIAQQ
'' SIG '' MA4wDAYKYIZIAWUDAgEDDTCBpwYDVR0fBIGfMIGcMDCg
'' SIG '' LqAshipodHRwOi8vY3JsLXNlcnZlci5vcmMuY29tL0NS
'' SIG '' THMvT1JDU1NQMy5jcmwwaKBmoGSGYmxkYXA6Ly9vcmMt
'' SIG '' ZHMub3JjLmNvbS9jbiUzZE9SQyUyMFNTUCUyMDMlMmNv
'' SIG '' JTNkT1JDJTIwUEtJJTJjYyUzZFVTP2NlcnRpZmljYXRl
'' SIG '' UmV2b2NhdGlvbkxpc3Q7YmluYXJ5MFkGA1UdEQRSMFCg
'' SIG '' JwYIYIZIAWUDBgagGwQZ00QQ2aiobIDUQa2haFghCELS
'' SIG '' ICiDRBDX5KAlBgorBgEEAYI3FAIDoBcMFUNoYXBlbGwu
'' SIG '' SmFzb25AZXBhLmdvdjAQBglghkgBZQMGCQEEAwEB/zAd
'' SIG '' BgNVHQ4EFgQUyYXNxgOnforpBKQslUZ2AdW4I/4wDQYJ
'' SIG '' KoZIhvcNAQELBQADggEBAIf8z+4QqlTaWoxnHlRf2Cp4
'' SIG '' j406iQghEqh//0ZgtivIWMSpQz5g876eGDKzfUIDOkVK
'' SIG '' nsjQ1KFr1fiMhM/rhdQsj3n+w+9+CBxsLomT3xVE3rz9
'' SIG '' 0DGYYPv/0F8MSUq17fLJ1aHY22lLf+MAi4XPP2bW7iOG
'' SIG '' yns1XhL2RH3IBzot6A9jQP/YM1o18cnnwhuhLMB9q+wg
'' SIG '' qnjeBqvWlbw3iuEzLC5fEbB4GL0uMBajrxNw4l7PAHnU
'' SIG '' Va0tbT47+YfYWWSK2kKvVUgg7L7kEM2PpFG2H98H+J2U
'' SIG '' UbDwkfueRCLP3ZMQCBE0WxkoHfxGsTLThgaAAHv9xwnF
'' SIG '' axa6y7RiDekwggY0MIIFHKADAgECAgICwjANBgkqhkiG
'' SIG '' 9w0BAQsFADBZMQswCQYDVQQGEwJVUzEYMBYGA1UEChMP
'' SIG '' VS5TLiBHb3Zlcm5tZW50MQ0wCwYDVQQLEwRGUEtJMSEw
'' SIG '' HwYDVQQDExhGZWRlcmFsIENvbW1vbiBQb2xpY3kgQ0Ew
'' SIG '' HhcNMTEwMTEyMDA1NDU3WhcNMjEwMTEyMDA1MjU5WjAz
'' SIG '' MQswCQYDVQQGEwJVUzEQMA4GA1UEChMHT1JDIFBLSTES
'' SIG '' MBAGA1UEAxMJT1JDIFNTUCAzMIIBIjANBgkqhkiG9w0B
'' SIG '' AQEFAAOCAQ8AMIIBCgKCAQEAp5bB2ZIV2oqyRvgyOhsX
'' SIG '' t5bdYoHhdAD3dBmi0HnUSkl3AdkqJUxZdmT18talN9v3
'' SIG '' QEadlP2BzX976pWEFGwA0bEMxMtJZfZq4I8At89H+5hP
'' SIG '' S6+s7xqfYeqDWNhCFREDl5gMGUBelpwNIvLI27U8D5tF
'' SIG '' btPBhBbc00zqyQCFpQ92IhbC57NF7K29/3cZE8i+VGfC
'' SIG '' 2PYxkWe/va+g7q5urEIUggGf6xt004uAtk7K+2Z43MKL
'' SIG '' eIyXIcuXH/bS50os2EIW62kwMkuKTM2XOJ9W7SvKbfGa
'' SIG '' O/tQxysErif+MX3FjyzdO8n27BR7mw15ZnBJLxnsY0Zc
'' SIG '' 8QbpXQG+MljtwwIDAQABo4IDKjCCAyYwDwYDVR0TAQH/
'' SIG '' BAUwAwEB/zBPBgNVHSAESDBGMAwGCmCGSAFlAwIBAwYw
'' SIG '' DAYKYIZIAWUDAgEDBzAMBgpghkgBZQMCAQMIMAwGCmCG
'' SIG '' SAFlAwIBAw0wDAYKYIZIAWUDAgEDETCB6QYIKwYBBQUH
'' SIG '' AQEEgdwwgdkwPwYIKwYBBQUHMAKGM2h0dHA6Ly9odHRw
'' SIG '' LmZwa2kuZ292L2ZjcGNhL2NhQ2VydHNJc3N1ZWRUb2Zj
'' SIG '' cGNhLnA3YzCBlQYIKwYBBQUHMAKGgYhsZGFwOi8vbGRh
'' SIG '' cC5mcGtpLmdvdi9jbj1GZWRlcmFsJTIwQ29tbW9uJTIw
'' SIG '' UG9saWN5JTIwQ0Esb3U9RlBLSSxvPVUuUy4lMjBHb3Zl
'' SIG '' cm5tZW50LGM9VVM/Y0FDZXJ0aWZpY2F0ZTtiaW5hcnks
'' SIG '' Y3Jvc3NDZXJ0aWZpY2F0ZVBhaXI7YmluYXJ5MIHKBggr
'' SIG '' BgEFBQcBCwSBvTCBujA4BggrBgEFBQcwBYYsaHR0cDov
'' SIG '' L2NybHNlcnZlci5vcmMuY29tL2NhQ2VydHMvT1JDU1NQ
'' SIG '' My5wN2MwfgYIKwYBBQUHMAWGcmxkYXA6Ly9vcmMtZHMu
'' SIG '' b3JjLmNvbS9jbiUzZE9SQyUyMFNTUCUyMDMlMmNPJTNk
'' SIG '' T1JDJTIwUEtJJTJjQyUzZFVTP2NBQ2VydGlmaWNhdGU7
'' SIG '' YmluYXJ5LGNyb3NzQ2VydGlmaWNhdGVQYWlyO2JpbmFy
'' SIG '' eTAOBgNVHQ8BAf8EBAMCAcYwHwYDVR0jBBgwFoAUrQx6
'' SIG '' dVzl85jEeZgOrCj9l/TnAvwwgbgGA1UdHwSBsDCBrTAq
'' SIG '' oCigJoYkaHR0cDovL2h0dHAuZnBraS5nb3YvZmNwY2Ev
'' SIG '' ZmNwY2EuY3JsMH+gfaB7hnlsZGFwOi8vbGRhcC5mcGtp
'' SIG '' Lmdvdi9jbiUzZEZlZGVyYWwlMjBDb21tb24lMjBQb2xp
'' SIG '' Y3klMjBDQSxvdSUzZEZQS0ksbyUzZFUuUy4lMjBHb3Zl
'' SIG '' cm5tZW50LGMlM2RVUz9jZXJ0aWZpY2F0ZVJldm9jYXRp
'' SIG '' b25MaXN0MB0GA1UdDgQWBBRqiG1S/LFEnjCuMxhNwDmd
'' SIG '' lmsksDANBgkqhkiG9w0BAQsFAAOCAQEAOARNKPrFvZOx
'' SIG '' nHFhuesSw106zf2i3IIdsB1UERlN0RsiBbwpqOoQemhV
'' SIG '' A9Wl6+KSj7IECVW6VMu7gs5F9+DoaidQzav/mQmQlTi2
'' SIG '' i8QYnPrCKn/kYeps5dnyaPbPyTNMRQoJsd4HvOktRGWE
'' SIG '' eDQrw/B5PtovkCpHLkG/2yFuk0eRg0XRZuIyOCjfhL1o
'' SIG '' oAQuqbWtNxU23XyX1GSAwdo+5wRm3OQYhgPOmEbOZCBY
'' SIG '' dqM/adoExzqCSqcUQn2DUMEdonQcIrddjI4DZB2oCu5x
'' SIG '' zfsPxh+B0RBlxwcwq4kP2ZuNsMRO4IrYZ1G+oEtO2drF
'' SIG '' 8e0xc7uOt5fFI1L/JnXkVDGCAcIwggG+AgEBMEgwMzEL
'' SIG '' MAkGA1UEBhMCVVMxEDAOBgNVBAoTB09SQyBQS0kxEjAQ
'' SIG '' BgNVBAMTCU9SQyBTU1AgMwIRAKre8vwBbNKdvVNlITHp
'' SIG '' epgwDAYIKoZIhvcNAgUFAKBOMBAGCisGAQQBgjcCAQwx
'' SIG '' AjAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMB8G
'' SIG '' CSqGSIb3DQEJBDESBBBxoXZUANr97Hp0yueic+4AMA0G
'' SIG '' CSqGSIb3DQEBAQUABIIBAHRd3GeU+BUXctWwQj6niujG
'' SIG '' 9mYRd3En+HHdt5hefLZ3tupGqvyoa07otKX+50Z0Wbg7
'' SIG '' I/qKh9XCkITqvigpCsi/Dzu38B6tpFdyPqaytKhqmA7B
'' SIG '' dVTQBUj829UaWrADIFQgJz/WRnE9i8fB2kXyCYAk6Xph
'' SIG '' 8gLnZiJ1uP1W8ur4VibpLLKIngSav+M0KNWV5ygoVRHb
'' SIG '' 0nonSYv/P6zexqZMP4gqo0lc+bJAlQWz6h21WPJTaJ9P
'' SIG '' hz2tPf301eQB5ydHB6/dryEoeO1iUt8iNMomchPwV9KC
'' SIG '' Kvu/islq0s0wehTOFLyFcBLlSS/VSediOpnBELy0H0ZV
'' SIG '' cCiv97J0yH0=
'' SIG '' End signature block
