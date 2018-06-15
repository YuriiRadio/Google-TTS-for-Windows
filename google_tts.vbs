'*********************************************************
'Google TTS for Windows on VBScript
'yurii.radio@gmail.com
'Yurii Radio - 2017
'*********************************************************

'You can run the script with a command-line option:
'	/lang:uk 	- languge
'	/utf8:false	- disable utf8 encode
'	/clipboard	- read text from clipboard
'	/cache - caching

'12.08.2017 fix - strScriptPath
'23.11.2017 add parameter - /utf8 (disable utf8 encode)
'28.11.2017 add parameter - /clipboard (read text from clipboard)
'15.06.2018 add parameter - /cache (caching)

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
Dim bClipboard
Dim bCache

Set objFSO = CreateObject("Scripting.FileSystemObject")

'Initialization section
strLang = "uk" 						'Default language
bUtf8 = True 						'Default UTF8 Encode
bCache = False						'Default no caching
bClipboard = False					'Default no read clipboard
strInputText = "Синтезатор Google" 	'Default Text
strMp3FileName = "response.mp3" 	'Default mp3 file
strPlayPrg = "madplay.exe"			'Console program for mp3 playback. Must be in the script directory
strScriptPath = objFSO.GetFile(WScript.ScriptFullName).ParentFolder 'Script Path
strSaveDir = objFSO.GetSpecialFolder(2)								'Save Dir (%Temp%)
'End Initialization

'MsgBox strInputText: WScript.Quit

Set objArgs = WScript.Arguments
If objArgs.Named.Exists("lang") Then
	strLang = objArgs.Named("lang")
End If

If objArgs.Named.Exists("utf8") And objArgs.Named("utf8") = "false" Then
	bUtf8 = False
End If

' Get clipboard text
If objArgs.Named.Exists("clipboard") Then
	bClipboard = True
	Dim objHTML
	Set objHTML = CreateObject("htmlfile")
	strInputText = objHTML.ParentWindow.ClipboardData.GetData("text")
End If

If objArgs.Unnamed.Count > 0 Then
	If objArgs.Unnamed(0) <> "" Then
		If Not bClipboard Then
			strInputText = ""
		End If
		Dim i: For i = 0 To objArgs.Unnamed.Count - 1
			If (objArgs.Unnamed.Count - 1) > i Then
				strInputText = strInputText & objArgs.Unnamed(i) & Chr(32)
			else
				strInputText = strInputText & objArgs.Unnamed(i)
			End If
		Next
	Else
		bUtf8 = False
	End If
Else
	bUtf8 = False
End If

If bUtf8 Then
	strInputText = UTF8Encode(strInputText)
End If

If objArgs.Named.Exists("cache") Then
	bCache = True
	Dim objMd5
	Set objMd5 = new MD5
	strMp3FileName = objMd5.MD5(strInputText) & ".mp3"
End If

If bCache And objFSO.FileExists(strSaveDir & Chr(47) & strMp3FileName) Then
	
	Set objShell = WScript.CreateObject("WScript.Shell")
	strCommand = strScriptPath & Chr(47) & strPlayPrg & Chr(32) & strSaveDir & Chr(47) & strMp3FileName
	objShell.Run strCommand, 0, true
	
	Set objShell = Nothing
else
	strInputText = URLEncode(strInputText)

	strURL = "https://translate.google.com/translate_tts?ie=UTF-8&client=tw-ob&q=" & strInputText & "&tl=" & strLang

	Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
	'Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP.3.0")
	objXMLHTTP.Open "GET", strURL, False
	objXMLHTTP.Send ' max 200 chars
	'MsgBox objXMLHTTP.Status: WScript.Quit

	If objXMLHTTP.Status = 200 Then
		'Adodb.Stream  - Antivirus ???
		Set objStream = createobject(Chr(65) & Chr(100) & Chr(111) & Chr(100) & Chr(98) & ".Stream")
		objStream.type = 1
		objStream.open
		objStream.write objXMLHTTP.responseBody
		objStream.savetofile strSaveDir & Chr(47) & strMp3FileName, 2
		objStream.Close
		Set objStream = Nothing

		Set objShell = WScript.CreateObject("WScript.Shell")
		strCommand = strScriptPath & Chr(47) & strPlayPrg & Chr(32) & strSaveDir & Chr(47) & strMp3FileName
		objShell.Run strCommand, 0, true

		If Not bCache Then
			objFSO.DeleteFile(strSaveDir & Chr(47) & strMp3FileName)
		End if
		Set objShell = Nothing
	End If
	
	Set objXMLHTTP = Nothing
End If 'If caching

Set objFSO = Nothing
'End Main Program

'*****Functions*****
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

Class MD5
	Private BITS_TO_A_BYTE
	Private BYTES_TO_A_WORD
	Private BITS_TO_A_WORD
	Private m_lOnBits(30)
	Private m_l2Power(30)

	Private Sub Class_Initialize         

		BITS_TO_A_BYTE = 8 'Const
		BYTES_TO_A_WORD = 4 'Const
		BITS_TO_A_WORD = 32 'Const
		
		m_lOnBits(0) = CLng(1)
		m_lOnBits(1) = CLng(3)
		m_lOnBits(2) = CLng(7)
		m_lOnBits(3) = CLng(15)
		m_lOnBits(4) = CLng(31)
		m_lOnBits(5) = CLng(63)
		m_lOnBits(6) = CLng(127)
		m_lOnBits(7) = CLng(255)
		m_lOnBits(8) = CLng(511)
		m_lOnBits(9) = CLng(1023)
		m_lOnBits(10) = CLng(2047)
		m_lOnBits(11) = CLng(4095)
		m_lOnBits(12) = CLng(8191)
		m_lOnBits(13) = CLng(16383)
		m_lOnBits(14) = CLng(32767)
		m_lOnBits(15) = CLng(65535)
		m_lOnBits(16) = CLng(131071)
		m_lOnBits(17) = CLng(262143)
		m_lOnBits(18) = CLng(524287)
		m_lOnBits(19) = CLng(1048575)
		m_lOnBits(20) = CLng(2097151)
		m_lOnBits(21) = CLng(4194303)
		m_lOnBits(22) = CLng(8388607)
		m_lOnBits(23) = CLng(16777215)
		m_lOnBits(24) = CLng(33554431)
		m_lOnBits(25) = CLng(67108863)
		m_lOnBits(26) = CLng(134217727)
		m_lOnBits(27) = CLng(268435455)
		m_lOnBits(28) = CLng(536870911)
		m_lOnBits(29) = CLng(1073741823)
		m_lOnBits(30) = CLng(2147483647)
		m_l2Power(0) = CLng(1)
		m_l2Power(1) = CLng(2)
		m_l2Power(2) = CLng(4)
		m_l2Power(3) = CLng(8)
		m_l2Power(4) = CLng(16)
		m_l2Power(5) = CLng(32)
		m_l2Power(6) = CLng(64)
		m_l2Power(7) = CLng(128)
		m_l2Power(8) = CLng(256)
		m_l2Power(9) = CLng(512)
		m_l2Power(10) = CLng(1024)
		m_l2Power(11) = CLng(2048)
		m_l2Power(12) = CLng(4096)
		m_l2Power(13) = CLng(8192)
		m_l2Power(14) = CLng(16384)
		m_l2Power(15) = CLng(32768)
		m_l2Power(16) = CLng(65536)
		m_l2Power(17) = CLng(131072)
		m_l2Power(18) = CLng(262144)
		m_l2Power(19) = CLng(524288)
		m_l2Power(20) = CLng(1048576)
		m_l2Power(21) = CLng(2097152)
		m_l2Power(22) = CLng(4194304)
		m_l2Power(23) = CLng(8388608)
		m_l2Power(24) = CLng(16777216)
		m_l2Power(25) = CLng(33554432)
		m_l2Power(26) = CLng(67108864)
		m_l2Power(27) = CLng(134217728)
		m_l2Power(28) = CLng(268435456)
		m_l2Power(29) = CLng(536870912)
		m_l2Power(30) = CLng(1073741824)
	End Sub

	Private Function LShift(lValue, iShiftBits)
		
		If iShiftBits = 0 Then
			LShift = lValue
			
			Exit Function
			
		ElseIf iShiftBits = 31 Then
			
			If lValue And 1 Then
				LShift = &H80000000
			Else
				LShift = 0
			End If
			
			Exit Function
			
		ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
			Err.Raise 6
		End If
		
		If (lValue And m_l2Power(31 - iShiftBits)) Then
			LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
		Else
			LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
		End If
		
	End Function

	Private Function RShift(lValue, iShiftBits)
		
		If iShiftBits = 0 Then
			RShift = lValue
			
			Exit Function
			
		ElseIf iShiftBits = 31 Then
			If lValue And &H80000000 Then
				RShift = 1
			Else
				RShift = 0
			End If
			
			Exit Function
			
		ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
			Err.Raise 6
		End If
		
		RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
		
		If (lValue And &H80000000) Then
			RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
		End If
		
	End Function

	Private Function RotateLeft(lValue, iShiftBits)
		RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
	End Function

	Private Function AddUnsigned(lX, lY)
	Dim lX4
	Dim lY4
	Dim lX8
	Dim lY8
	Dim lResult

		lX8 = lX And &H80000000
		lY8 = lY And &H80000000
		lX4 = lX And &H40000000
		lY4 = lY And &H40000000
		lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)

		If lX4 And lY4 Then
			lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
		ElseIf lX4 Or lY4 Then
			If lResult And &H40000000 Then
				lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
			Else
				lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
			End If
		Else
			lResult = lResult Xor lX8 Xor lY8
		End If
		
		AddUnsigned = lResult
		
	End Function

	Private Function F(x, y, z)
		F = (x And y) Or ((Not x) And z)
	End Function

	Private Function G(x, y, z)
		G = (x And z) Or (y And (Not z))
	End Function

	Private Function H(x, y, z)
		H = (x Xor y Xor z)
	End Function

	Private Function I(x, y, z)
		I = (y Xor (x Or (Not z)))
	End Function

	Private Sub FF(a, b, c, d, x, s, ac)
		a = AddUnsigned(a, AddUnsigned(AddUnsigned(F(b, c, d), x), ac))
		a = RotateLeft(a, s)
		a = AddUnsigned(a, b)
	End Sub

	Private Sub GG(a, b, c, d, x, s, ac)
		a = AddUnsigned(a, AddUnsigned(AddUnsigned(G(b, c, d), x), ac))
		a = RotateLeft(a, s)
		a = AddUnsigned(a, b)
	End Sub

	Private Sub HH(a, b, c, d, x, s, ac)
		a = AddUnsigned(a, AddUnsigned(AddUnsigned(H(b, c, d), x), ac))
		a = RotateLeft(a, s)
		a = AddUnsigned(a, b)
	End Sub

	Private Sub II(a, b, c, d, x, s, ac)
		a = AddUnsigned(a, AddUnsigned(AddUnsigned(I(b, c, d), x), ac))
		a = RotateLeft(a, s)
		a = AddUnsigned(a, b)
	End Sub

	Private Function ConvertToWordArray(sMessage)
	Dim lMessageLength
	Dim lNumberOfWords
	Dim lWordArray()
	Dim lBytePosition
	Dim lByteCount
	Dim lWordCount
	Const MODULUS_BITS = 512
	Const CONGRUENT_BITS = 448

		lMessageLength = Len(sMessage)
		lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
		ReDim lWordArray(lNumberOfWords - 1)
		lBytePosition = 0
		lByteCount = 0

		Do Until lByteCount >= lMessageLength
			lWordCount = lByteCount \ BYTES_TO_A_WORD
			lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
			lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
			lByteCount = lByteCount + 1
		Loop
		
		lWordCount = lByteCount \ BYTES_TO_A_WORD
		lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
		lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)
		lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
		lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
		
		ConvertToWordArray = lWordArray
		
	End Function

	Private Function WordToHex(lValue)
	Dim lByte
	Dim lCount

		For lCount = 0 To 3
			lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
			WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
		Next
		
	End Function

	Public Function MD5(sMessage)
	Dim x
	Dim k
	Dim AA
	Dim BB
	Dim CC
	Dim DD
	Dim a
	Dim b
	Dim c
	Dim d
	Const S11 = 7
	Const S12 = 12
	Const S13 = 17
	Const S14 = 22
	Const S21 = 5
	Const S22 = 9
	Const S23 = 14
	Const S24 = 20
	Const S31 = 4
	Const S32 = 11
	Const S33 = 16
	Const S34 = 23
	Const S41 = 6
	Const S42 = 10
	Const S43 = 15
	Const S44 = 21

		x = ConvertToWordArray(sMessage)
		a = &H67452301
		b = &HEFCDAB89
		c = &H98BADCFE
		d = &H10325476
		
		For k = 0 To UBound(x) Step 16
			AA = a
			BB = b
			CC = c
			DD = d
			FF a, b, c, d, x(k + 0), S11, &HD76AA478
			FF d, a, b, c, x(k + 1), S12, &HE8C7B756
			FF c, d, a, b, x(k + 2), S13, &H242070DB
			FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
			FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
			FF d, a, b, c, x(k + 5), S12, &H4787C62A
			FF c, d, a, b, x(k + 6), S13, &HA8304613
			FF b, c, d, a, x(k + 7), S14, &HFD469501
			FF a, b, c, d, x(k + 8), S11, &H698098D8
			FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
			FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
			FF b, c, d, a, x(k + 11), S14, &H895CD7BE
			FF a, b, c, d, x(k + 12), S11, &H6B901122
			FF d, a, b, c, x(k + 13), S12, &HFD987193
			FF c, d, a, b, x(k + 14), S13, &HA679438E
			FF b, c, d, a, x(k + 15), S14, &H49B40821
			GG a, b, c, d, x(k + 1), S21, &HF61E2562
			GG d, a, b, c, x(k + 6), S22, &HC040B340
			GG c, d, a, b, x(k + 11), S23, &H265E5A51
			GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
			GG a, b, c, d, x(k + 5), S21, &HD62F105D
			GG d, a, b, c, x(k + 10), S22, &H2441453
			GG c, d, a, b, x(k + 15), S23, &HD8A1E681
			GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
			GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
			GG d, a, b, c, x(k + 14), S22, &HC33707D6
			GG c, d, a, b, x(k + 3), S23, &HF4D50D87
			GG b, c, d, a, x(k + 8), S24, &H455A14ED
			GG a, b, c, d, x(k + 13), S21, &HA9E3E905
			GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
			GG c, d, a, b, x(k + 7), S23, &H676F02D9
			GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
			HH a, b, c, d, x(k + 5), S31, &HFFFA3942
			HH d, a, b, c, x(k + 8), S32, &H8771F681
			HH c, d, a, b, x(k + 11), S33, &H6D9D6122
			HH b, c, d, a, x(k + 14), S34, &HFDE5380C
			HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
			HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
			HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
			HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
			HH a, b, c, d, x(k + 13), S31, &H289B7EC6
			HH d, a, b, c, x(k + 0), S32, &HEAA127FA
			HH c, d, a, b, x(k + 3), S33, &HD4EF3085
			HH b, c, d, a, x(k + 6), S34, &H4881D05
			HH a, b, c, d, x(k + 9), S31, &HD9D4D039
			HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
			HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
			HH b, c, d, a, x(k + 2), S34, &HC4AC5665
			II a, b, c, d, x(k + 0), S41, &HF4292244
			II d, a, b, c, x(k + 7), S42, &H432AFF97
			II c, d, a, b, x(k + 14), S43, &HAB9423A7
			II b, c, d, a, x(k + 5), S44, &HFC93A039
			II a, b, c, d, x(k + 12), S41, &H655B59C3
			II d, a, b, c, x(k + 3), S42, &H8F0CCC92
			II c, d, a, b, x(k + 10), S43, &HFFEFF47D
			II b, c, d, a, x(k + 1), S44, &H85845DD1
			II a, b, c, d, x(k + 8), S41, &H6FA87E4F
			II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
			II c, d, a, b, x(k + 6), S43, &HA3014314
			II b, c, d, a, x(k + 13), S44, &H4E0811A1
			II a, b, c, d, x(k + 4), S41, &HF7537E82
			II d, a, b, c, x(k + 11), S42, &HBD3AF235
			II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
			II b, c, d, a, x(k + 9), S44, &HEB86D391
			a = AddUnsigned(a, AA)
			b = AddUnsigned(b, BB)
			c = AddUnsigned(c, CC)
			d = AddUnsigned(d, DD)
		Next

		MD5 = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
		
	End Function

End Class