'*********************************************************
'Google SAPI for Windows on VBScript
'yurii.radio@gmail.com
'Yurii Radio - 2017
'*********************************************************

'You can run the script with a command-line option:
'	/lang:uk
'	/utf8:false

'12.08.2017 fix - strScriptPath
'23.11.2017 add parameter - utf8 (default True)

Option Explicit

Dim strInputText
Dim strLang
Dim strMp3FileName
Dim strURL
Dim strSaveDir
Dim strScriptPath
Dim strCommand
Dim strPlayPrg

Dim objArgs
Dim objFSO
Dim objXMLHTTP
Dim objStream
Dim objShell

Dim bUtf8

Set objFSO = CreateObject("Scripting.FileSystemObject")

'Initialization section
strLang = "uk" 						'Default language
bUtf8 = True 						'Default UTF8 Encode (True, False)
strInputText = "Синтезатор Google" 	'Default Text
strMp3FileName = "response.mp3" 	'Default mp3 file
strPlayPrg = "madplay.exe"			'Console program for mp3 playback. Must be in the script directory
strScriptPath = objFSO.GetFile(WScript.ScriptFullName).ParentFolder 'Script Path
strSaveDir = objFSO.GetSpecialFolder(2)								'Save Dir (%Temp%)
'End Initialization


Set objArgs = WScript.Arguments
If objArgs.Named.Exists("lang") Then
	strLang = objArgs.Named("lang")
End If

If objArgs.Named.Exists("utf8") And objArgs.Named("utf8") = "false" Then
	bUtf8 = False
End If

If objArgs.Unnamed.Count > 0 Then
	If objArgs.Unnamed(0) <> "" Then
		strInputText = ""
		Dim i: For i = 0 To objArgs.Unnamed.Count - 1
			strInputText = strInputText & objArgs.Unnamed(i)
			If (objArgs.Unnamed.Count - 1) > i Then
				strInputText = strInputText & Chr(32)
			End If
		Next
		If bUtf8 Then
			strInputText = UTF8Encode(strInputText)
		End If
		strInputText = URLEncode(strInputText)
	End If
End If
'MsgBox strInputText: WScript.Quit

strURL = "https://translate.google.com/translate_tts?ie=UTF-8&client=tw-ob&q=" & strInputText & "&tl=" & strLang

Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
objXMLHTTP.Open "GET", strURL, False
objXMLHTTP.Send ' max 200 chars
'MsgBox objXMLHTTP.Status: WScript.Quit

If objXMLHTTP.Status = 200 Then
	Set objStream = createobject("Adodb.Stream")
	objStream.type = 1
	objStream.open
	objStream.write objXMLHTTP.responseBody
	objStream.savetofile strSaveDir & Chr(47) & strMp3FileName, 2
	objStream.Close
	Set objStream = Nothing

	Set objShell = WScript.CreateObject("WScript.Shell")
	strCommand = strScriptPath & Chr(47) & strPlayPrg & Chr(32) & strSaveDir & Chr(47) & strMp3FileName
	objShell.Run strCommand, 0, true

	objFSO.DeleteFile(strSaveDir & Chr(47) & strMp3FileName)
	Set objShell = Nothing
End If

Set objXMLHTTP = Nothing
Set objFSO = Nothing

Function UTF8Encode(s)
    Dim i, c, utfc, b1, b2, b3
    
    For i=1 to Len(s)
        c = ToLong(AscW(Mid(s,i,1)))
 
        If c < 128 Then
            utfc = chr(c)
        ElseIf c < 2048 Then
            b1 = c Mod &h40
            b2 = (c - b1) / &h40
            utfc = chr(&hC0 + b2) & chr(&h80 + b1)
        ElseIf c < 65536 And (c < 55296 Or c > 57343) Then
            b1 = c Mod &h40
            b2 = ((c - b1) / &h40) Mod &h40
            b3 = (c - b1 - (&h40 * b2)) / &h1000
            utfc = chr(&hE0 + b3) & chr(&h80 + b2) & chr(&h80 + b1)
        Else
            utfc = Chr(&hEF) & Chr(&hBF) & Chr(&hBD)
        End If

        UTF8Encode = UTF8Encode + utfc
    Next
End Function

Function ToLong(intVal)
    If intVal < 0 Then
        ToLong = CLng(intVal) + &H10000
    Else
        ToLong = CLng(intVal)
    End If
End Function

Function URLEncode(sStr)
    Dim i, acode

    URLEncode = sStr

    For i = Len(URLEncode) To 1 Step -1
        acode = Asc(Mid(URLEncode, i, 1))
        If (aCode >=48 And aCode <=57) Or (aCode >= 65 And aCode <=90) Or (aCode >= 97 And aCode <=122) Then
        ' don't touch alphanumeric chars
        Else
            Select Case acode
            Case 32
                ' replace space with "+"
                URLEncode = Left(URLEncode, i - 1) & "+" & Mid(URLEncode, i + 1)
            Case Else
                ' replace punctuation chars with "%hex"
                URLEncode = Left(URLEncode, i - 1) & "%" & Hex(acode) & Mid(URLEncode, i + 1)
            End Select
        End If
    Next
End Function