Option Explicit

Dim strInputText
Dim strLang
Dim strMp3FileName
Dim strURL
Dim strSaveDir
Dim strScriptPath
Dim strCommand

Dim objArgs
Dim objFSO
Dim objXMLHTTP
Dim objStream
Dim objShell

Set objFSO = CreateObject("Scripting.FileSystemObject")
strScriptPath = objFSO.GetAbsolutePathName(".\") & "\"
'strSaveDir = "D:\Temp\Tests\"
strSaveDir = objFSO.GetSpecialFolder(2) & "\"

Set objArgs = WScript.Arguments
If objArgs.Named.Exists("lang") Then
	strLang = objArgs.Named("lang")
else
	strLang = "uk"
End If

strInputText = "Синтезатор Google"
If objArgs.Unnamed.Count > 0 Then
	If objArgs.Unnamed(0) <> "" Then
		strInputText = ""
		Dim i: For i = 0 To objArgs.Unnamed.Count - 1
			strInputText = strInputText & objArgs.Unnamed(i)
			If (objArgs.Unnamed.Count - 1) > i Then strInputText = strInputText & Chr(32) End If
		Next
		strInputText = UTF8Encode(strInputText)
		strInputText = URLEncode(strInputText)
	End If
End If
'MsgBox strInputText: WScript.Quit

strMp3FileName = "response.mp3"
strURL = "https://translate.google.com/translate_tts?ie=UTF-8&client=tw-ob&q=" & strInputText & "&tl=" & strLang

Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
objXMLHTTP.Open "GET", strURL, False
objXMLHTTP.Send

Set objStream = createobject("Adodb.Stream")
objStream.type = 1
objStream.open
objStream.write objXMLHTTP.responseBody
objStream.savetofile strSaveDir & strMp3FileName, 2
objStream.Close
Set objStream = Nothing
Set objXMLHTTP = Nothing

Set objShell = WScript.CreateObject("WScript.Shell")
strCommand = "madplay.exe " & strSaveDir & strMp3FileName
objShell.Run strCommand, 0, true

'strCommand = scriptPath & "@del " & strMp3FileName
'objShell.Run strCommand, 0, false
objFSO.DeleteFile(strSaveDir & strMp3FileName)

Set objShell = Nothing
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

Function EncodeWithoutBOM (strText)
	Const adSaveCreateNotExist = 1
	Const adSaveCreateOverWrite = 2
	Const adTypeBinary = 1
	Const adTypeText   = 2

	Dim objStreamUTF8      : Set objStreamUTF8      = CreateObject("ADODB.Stream")
	Dim objStreamUTF8NoBOM : Set objStreamUTF8NoBOM = CreateObject("ADODB.Stream")

	With objStreamUTF8
	  .Open
	  .Type    = adTypeText
	  .Charset = "UTF-8"
	  .WriteText strText
	  .Position = 0
	  .SaveToFile "testUTF8.txt", adSaveCreateOverWrite
	  .Type     = adTypeBinary
	  .Position = 3
	End With

	With objStreamUTF8NoBOM
	  .Open
	  .Type  = adTypeBinary
	  objStreamUTF8.CopyTo objStreamUTF8NoBOM
	  .SaveToFile "testUTF8NoBOM.txt", adSaveCreateOverWrite
	  .Type    = adTypeText
	  EncodeWithoutBOM = .ReadText
	End With
	
	objStreamUTF8.Close
	objStreamUTF8NoBOM.Close
End Function