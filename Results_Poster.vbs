'''MD5'''
' Source http://chayoung.tistory.com/entry/VBScript-MD5
' found via http://stackoverflow.com/questions/10198690/how-to-generate-md5-using-vb-in-classic-asp/10198875#10198875

Private Const BITS_TO_A_BYTE = 8
Private Const BYTES_TO_A_WORD = 4
Private Const BITS_TO_A_WORD = 32
Private m_lOnBits(30)
Private m_l2Power(30)

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
'''MD5-END'''

Dim folderName, key, secretKey, strComputerName, hostFileName
Set objSuperFolder = Nothing
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set wshNetwork = WScript.CreateObject( "WScript.Network" )
'Wscript.Echo "Poster started"

' Get path to post folder, and get PractiTest keys
If GetParamsFromIniFile(out_FolderName, out_Key, out_Secret_Key) = True Then
	folderName = out_FolderName
	key = out_Key
	secretKey = out_Secret_Key
	Set objSuperFolder = objFSO.GetFolder(folderName)
	strComputerName = wshNetwork.ComputerName
	hostFileName = objFSO.BuildPath(folderName, strComputerName & ".txt")
Else
	Wscript.Echo "Poster Stopped, can not read from INI file"
	WScript.Quit 0
End If
Do
	' Write to HostName.txt file - Indicates which host is running
	Call WriteLineToFile(hostFileName, Now() & ": " & Wscript.ScriptFullName)
	If PickFileFromFolder(objSuperFolder, "STOP", 0, out_ObjXmlFile) Then
		'Call objFSO.DeleteFile(objFSO.GetFile(hostFileName), True)	
		'Wscript.Echo "Poster Stopped"
		WScript.Quit 0
	End If

	If PickXmlFileFromFolder(objSuperFolder, out_ObjXmlFile) Then
		If ProcessXmlFile(out_ObjXmlFile, out_NewObjXmlFile) Then
			Call MoveXmlFile("Posted", out_NewObjXmlFile)
		Else
			Call MoveXmlFile("Error", out_NewObjXmlFile)
		End If
	Else
		WScript.Sleep 10000
	End If
	
Loop While True


Function GetParamsFromIniFile(out_FolderName, out_Key, out_Secret_Key)
	out_FolderName = ""
	out_Key = ""
	out_Secret_Key = ""
	
	strPath = Wscript.ScriptFullName
	strPath = Replace(strPath, ".vbs", ".ini")
	
	out_FolderName = ReadIni(strPath, "General", "folderName ")
	out_Key = ReadIni(strPath, "Practitest", "key ")
	out_Secret_Key = ReadIni(strPath, "Practitest", "secretKey ")
	
	If out_FolderName = "" OR out_Key = "" OR out_Secret_Key = "" Then
		GetParamsFromIniFile = False
		Wscript.Echo "ERROR: out_FolderName = " & out_FolderName & "; out_Key = " & out_Key & "; out_Secret_Key = " & out_Secret_Key
	Else
		GetParamsFromIniFile = True
	End If
End Function

Function GetXmlPropAndValue(line, out_Prop, out_Value)
	GetXmlPropAndValue = False
	Set regEx = New RegExp
	regEx.Global = True
	regEx.Pattern = "([^<>]*)"
	Set matches = regEx.Execute(line)
	If matches.Count = 9 Then
		out_Prop = matches(2).value
		out_Value = matches(4).value
		GetXmlPropAndValue = True
	End If
	Set regEx = Nothing
	Set matches = Nothing
End Function


Function GetRegexFirstMatch(str, pattern)
	GetRegexFirstMatch = ""
    If NOT str = "" Then
      Set regEx = New RegExp
      regEx.Global = True
      regEx.Pattern = pattern
      Set matches = regEx.Execute(str)
      If matches.Count > 0 Then
		 GetRegexFirstMatch = matches(0).value
	  End If
	End If
	Set regEx = Nothing 
	Set matches = Nothing
End Function


Function WriteLineToFile(hostFileName, txtContent)
	If objFSO.FileExists(hostFileName) Then
	  Set spoFile = objFSO.OpenTextFile(hostFileName, 2, True)
	Else
	  Set spoFile = objFSO.CreateTextFile(hostFileName, True)
	End If
	
	spoFile.Write txtContent
	spoFile.Close
	Set spoFile = Nothing
End Function


Function ProcessXmlFile(objXmlFile, out_ObjXmlFile)
	' Lock file - Rename file before using
	lockFileName = objXmlFile & "." & strComputerName
	objXmlFile.Move(lockFileName)
	Set out_ObjXmlFile = objXmlFile
	ProcessXmlFile = PostXmlFileContent(lockFileName)
End Function


Function PostXmlFileContent(xmlFileName)
	PostXmlFileContent = False
	ReDim arrProp(1)
	ReDim arrValue(1)
	' Open the file for input.
	Set MyFile = objFSO.OpenTextFile(xmlFileName, 1, True)

	' Read from the file and display the results.
	Do While MyFile.AtEndOfStream <> True
		line = MyFile.ReadLine
		If GetXmlPropAndValue(line, out_Prop, out_Value) Then
			arrProp(UBound(arrProp) - 1) = out_Prop
			arrValue(UBound(arrValue) - 1) = out_Value
			ReDim Preserve arrProp(UBound(arrProp) + 1)
			ReDim Preserve arrValue(UBound(arrValue) + 1)
		End If
	Loop
	MyFile.Close
	ReDim Preserve arrProp(UBound(arrProp) - 2)
	ReDim Preserve arrValue(UBound(arrValue) - 2)
	
	Set MyFile = objFSO.OpenTextFile(xmlFileName, 8, True)
	
	' Post JSON to PractiTest
	If PostToPractiTest(arrProp, arrValue, out_Status, out_StatusText, out_ResponseText, out_JsonString) Then
		
		MyFile.WriteLine vbLf & "***LOG***" & Now() & "***: Post was successful; Status = " & out_Status & "; StatusText = " & out_StatusText
		MyFile.WriteLine vbLf & "***LOG***" & Now() & "***: JSON STRING: " & out_JsonString
		PostXmlFileContent = True
	Else
		MyFile.WriteLine "***ERROR***" & Now() & "***: Status = " & out_Status & "; StatusText = " & out_StatusText 
		MyFile.WriteLine "***ERROR***" & Now() & "***: JSON STRING: " & out_JsonString
		MyFile.WriteLine "ResponseText = " & out_ResponseText
	End If
	
	MyFile.Close
	Set MyFile = Nothing
End Function


Function PostToPractiTest(propArray, valueArray, out_Status, out_StatusText, out_ResponseText, out_JsonString)
	PostToPractiTest = False
	ts = datediff("s",#1970/1/1#,now())

	URL = "https://prod.practitest.com/api/automated_tests/upload_test_result.json"
	signature = GetStringCheckSum(key & secretKey & ts)
	authStr = "custom api_key=" & key & ", signature=" & signature & ", ts=" & ts

	propValueStr = ""
	For w = LBound(propArray) to UBound(propArray)
		If propArray(w) = "project_id" Or propArray(w) = "testset_display_id" Or propArray(w) = "test_display_id" Or propArray(w) = "exit_code" Then
			propValueStr = propValueStr & """" & propArray(w) & """: """ & valueArray(w) & """, "
			PostToPractiTest = True
		ElseIf propArray(w) = "instance_custom_fields" Then
			propValueStr = propValueStr & """" & propArray(w) & """:" & valueArray(w) & ", "
			PostToPractiTest = True
		End If
	Next
	
	If PostToPractiTest = True Then
		PostToPractiTest = False
		' Remove ", " from the end of the string
		propValueStr = Left(propValueStr, Len(propValueStr) - 2)
		
		'strJSONToSend = "{""project_id"": ""1328"", ""testset_display_id"": ""99"", ""test_display_id"": ""320"", ""exit_code"": ""0""}"
		strJSONToSend = "{" & propValueStr & "}"
		out_JsonString = strJSONToSend
		Set objRequest = CreateObject("MSXML2.ServerXMLHTTP")
		objRequest.open "POST", URL, False 
		objRequest.setRequestHeader "Authorization", authStr 
		objRequest.setRequestHeader "Content-Type", "application/json" 

		objRequest.send strJSONToSend
		
		out_Status = objRequest.status
		out_StatusText = objRequest.statusText
		out_ResponseText = objRequest.responseText
		set objRequest = nothing
		If out_StatusText = "OK" Then
			PostToPractiTest = True
		End If
	End If
End Function


Function MoveXmlFile(subFolder, new_ObjXmlFile)
	newPath = objFSO.BuildPath(folderName, subFolder)
	Call CreateFolderIfNotExist(newPath)
	newSubPostedPath = objFSO.BuildPath(newPath, Year(Now()) & "-" & Right("0" & Month(Now()),2) & "-" & Right("0" & Day(Now()),2))
	Call CreateFolderIfNotExist(newSubPostedPath)
	' Rename file - add '.xml' to file name
	newFilePath = objFSO.BuildPath(newSubPostedPath, new_ObjXmlFile.name & ".xml")
	Do While objFSO.FileExists(newFilePath)
		newFilePath = newFilePath & "1"
	Loop
	' Move to 'posted' > 'today date' folder
	new_ObjXmlFile.Move(newFilePath)
	
	Set new_ObjXmlFile = Nothing
End Function


Function GetBuildPath(path)
   Dim newpath
   newpath = objFSO.BuildPath(path, "Sub Folder") 
   GetBuildPath = newpath
End Function


Function PickXmlFileFromFolder(fFolder, out_ObjXmlFile)
	fileExtension = "XML"
	PickXmlFileFromFolder = PickFileFromFolder(fFolder, fileExtension, 10, out_ObjXmlFile)
End Function


Function PickFileFromFolder(fFolder, fileExtension, minFileSize, out_ObjXmlFile)
	Set out_ObjXmlFile = Nothing
    Set objFolder = objFSO.GetFolder(fFolder.Path)
    Set colFiles = objFolder.Files
    For Each objFile in colFiles
        If UCase(objFSO.GetExtensionName(objFile.name)) = UCase(fileExtension) Then
			If objFile.size >= minFileSize Then
				If (out_ObjXmlFile is Nothing) Then
					Set out_ObjXmlFile = objFile
					PickFileFromFolder = True
				ElseIf (objFile.DateLastModified < out_ObjXmlFile.DateLastModified) Then
					Set out_ObjXmlFile = objFile
				End If
			Else
				Call MoveXmlFile("Error", objFile)
			End If
        End If
    Next
    Set objFolder = Nothing
    Set colFiles = Nothing
End Function


Function CreateFolderIfNotExist(path)
  Set fileSys = CreateObject("Scripting.FileSystemObject") 
  If Not fileSys.FolderExists(path) Then 
    fileSys.CreateFolder(path) 
  End If
  Set fileSys = Nothing
End Function


Function ReadIni( myFilePath, mySection, myKey )
    ' This function returns a value read from an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be returned
    '
    ' Returns:
    ' the [string] value for the specified key in the specified section
    '
    ' CAVEAT:     Will return a space if key exists but value is blank
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre and Rob van der Woude

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection

    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ReadIni     = ""
    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )

    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, ForReading, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )

            ' Check if section is found in the current line
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                strLine = Trim( objIniFile.ReadLine )

                ' Parse lines until the next section is reached
                Do While Left( strLine, 1 ) <> "["
                    ' Find position of equal sign in the line
                    intEqualPos = InStr( 1, strLine, "=", 1 )
                    If intEqualPos > 0 Then
                        strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                        ' Check if item is found in the current line
                        If LCase( strLeftString ) = LCase( strKey ) Then
                            ReadIni = Trim( Mid( strLine, intEqualPos + 1 ) )
                            ' In case the item exists but value is blank
                            If ReadIni = "" Then
                                ReadIni = " "
                            End If
                            ' Abort loop when item is found
                            Exit Do
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = Trim( objIniFile.ReadLine )
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
        WScript.Echo strFilePath & " doesn't exists. Exiting..."
        Wscript.Quit 1
    End If
	Set objIniFile = Nothing
End Function


'''MD5-CODE-START'''
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

Public Function GetStringCheckSum(sMessage)
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

	GetStringCheckSum = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
End Function
'''MD5-CODE-END''''