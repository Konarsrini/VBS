'#=====================================
Dim Arg,strProcess
Set Arg = WScript.Arguments
strMethod = Arg(0)
strURL = Arg(1)
strUN = Arg(2)
strPW = Arg(3)
strCon = Arg(4)
'msgbox strCon
'#=====================================
WebService_GET strMethod,strURL,strUN,strPW,strCon
strFilePathWithName = "G:\VBS\VB_LibraryFunctions\GivePopUpMessage.vbs"
strText = "...Completed..refer G:\API\ServerResponse.txt for log"
strDuration = "3"
strTitle= "Fading Message..."
strVBSFilePathArgs = """"&strFilePathWithName&""" """&strText&""" """&strDuration&""" """&strTitle&""""
CallLocalVBS(strVBSFilePathArgs)
'=====================================
'Msxml2.ServerXMLHTTP: This works ok for GET and POST only (PUT and DELETE - check and proceed)
Public Function WebService_GET(strMethod,strURL,strUN,strPW,strCon)
	WebService_GET = ""
	strCon = Replace(strCon,"8$9",chr(34))
	'msgbox strCon
	Set objXmlHttpMain = CreateObject("Msxml2.ServerXMLHTTP.6.0") 
	objXmlHttpMain.open strMethod,strURL,True,strUN,strPW
	objXmlHttpMain.setRequestHeader "Content-Type", "application/json"
	objXmlHttpMain.setRequestHeader "Accept", "*/*"
	'msgbox strCon
	objXmlHttpMain.send strCon
	strResText = objXmlHttpMain.responseText
	strHttpStat = objXmlHttpMain.status
	strResHeaderinfo = objXmlHttpMain.getAllResponseHeaders()
 	WebService_GET = VbCrLF&"=========="&VbCrLF&strResHeaderinfo&VbCrLF&"=========="&VbCrLF&strResText&VbCrLF&"=========="&VbCrLF
	strDesc = WebService_GET
	strtxtFilePathWithName = "G:\API\ServerResponse.txt"
	strFilePathWithName = "G:\VBS\VB_LibraryFunctions\MyMcLogTXT1.vbs"
	strDescA2 = "====================END==========================="
	'--------------
	strDescA1 = "====BaseURL: "&strURL&"====="&strMethod&"  method=====HttpStatus: "&strHttpStat&"===="&date&"="&Time&"===="	
	'------------------
	strVBSFilePathArgs = ""&strFilePathWithName&" "&strtxtFilePathWithName&" "&strDescA2
	CallLocalVBS(strVBSFilePathArgs)
	'------------------
	strVBSFilePathArgs = ""&strFilePathWithName&" "&strtxtFilePathWithName&" "&strResText
	CallLocalVBS(strVBSFilePathArgs)
	'------------------
	strVBSFilePathArgs1 = ""&strFilePathWithName&" "&strtxtFilePathWithName&" """&strDescA1&VbCrLF&strResHeaderinfo&VbCrLF&""""	
	CallLocalVBS(strVBSFilePathArgs1)
	'--------------
End Function
'***********************************************************
Public Function CallLocalVBS(strVBSFilePathArgs)
	Set WSHShell = CreateObject("WScript.Shell")	
	'WSHShell.Run "wscript """&strVBSFilePath&""" """&strArgs1&"""", , True
	'WSHShell.Run(strVBSFilePath&" "&strArgs2&" "&strArgs3)
	WSHShell.Run(strVBSFilePathArgs)
End Function
'=============
'***********************************************************
