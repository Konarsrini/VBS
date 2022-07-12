Dim Arg,strProcess
Set Arg = WScript.Arguments
strFilePathWithName = Arg(0)
strDescText = Arg(1)
'msgbox "A: "&Arg(0)
'msgbox "Desc: "&Arg(1)
MyMcLogTXT1 strFilePathWithName&"8$9"&strDescText
''**********************************************************************
Public Function MyMcLogTXT1(strFilePathWithName89strDesc)
	arrFilePathWithName89strDesc = Split(strFilePathWithName89strDesc,"8$9")
	strFilePathWithName = arrFilePathWithName89strDesc(0)
	strDescText = Replace(strFilePathWithName89strDesc,strFilePathWithName&"8$9","")
	 'Open in utf-8 , append and Save file in utf8 using ADODB.Stream object
	 Dim fso,ObjNoteOpen,ObjCreateText
	 Dim Text,objectselection 
	 Dim adoStream,adoStreamOut 
	 'msgbox strFilePathWithName
	 'msgbox strDescText
	 Set fso = CreateObject("Scripting.FileSystemObject")
	 If Not fso.FileExists(strFilePathWithName) Then
	    Set ObjCreateText = fso.CreateTextFile(strFilePathWithName)
	    Set ObjCreateText = Nothing
	 End If
	 Set adoStream = CreateObject("ADODB.Stream")
	 adoStream.Charset = "UTF-8"
	 adoStream.Open
	 '---
	 'On Error Resume Next
	 strTime2 = Time
	 Do 
	 	Err.Clear
		'On Error Resume Next
	 	'Wait 0,100
		adoStream.LoadFromFile strFilePathWithName
	 	strErrNum = Err.Number
	 	'Wait 0,100
		'On Error Goto 0
	 Loop Until strErrNum = 0 or DateDiff("s",strTime2,Time) > 10
	 On Error Goto 0
	 strG = adoStream.ReadText
	 '---
	 Set adoStreamOut = CreateObject("ADODB.Stream")
	 adoStreamOut.Charset = "UTF-8"
	 adoStreamOut.Open
	 '---------
	 'On Error Resume Next
	 strTime1 = Time
	 Do 
	 	Err.Clear
		'On Error Resume Next
	 	'Wait 0,100
		'msgbox adoStreamOut.read()
		adoStreamOut.WriteText strDescText&VbCrLF&strG
	 	strErrNum = Err.Number
	 	'Wait 0,100
		On Error Goto 0
	 Loop Until strErrNum = 0 or DateDiff("s",strTime1,Time) > 10
	 On Error Goto 0
	 '------------
	 '---------
	 strTime1 = Time
	 Do 
	 	Err.Clear
		'On Error Resume Next
	 	'Wait 0,100
	 	adoStreamOut.SaveToFile strFilePathWithName, 2
	 	strErrNum = Err.Number
	 	'Wait 0,100
		'On Error Goto 0
	 Loop Until strErrNum = 0 or DateDiff("s",strTime1,Time) > 10 
	 On Error Goto 0
	 '------------
	 'adoStreamOut.SaveToFile strFilePathWithName, 2
	 ''1: Creates a new file if the file does not already exist
	 ''2: Overwrites the file with the data from the currently open Stream object, if the file already exists 
		'1 is  no Set fso = Nothing
	 Set adoStreamOut = Nothing 
	 
	 Set adoStream = Nothing
 On Error Goto 0

End Function
'******************