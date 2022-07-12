Dim Arg,strProcess
Set Arg = WScript.Arguments
strText = Arg(0)
intDurationSec = Arg(1)
strTitle = Arg(2)
'msgbox strText
'msgbox intDurationSec
'msgbox strTitle
'#=====================================
GivePopUpMessage strText,intDurationSec,strTitle
 '**********************************************************************
 Public Function GivePopUpMessage(strText,intDurationSec,strTitle)
	'msgbox strText
	'msgbox intDurationSec
	'msgbox strTitle
 	Dim objShell
	Set objShell = CreateObject("Wscript.Shell")
	objShell.Popup strText,intDurationSec,strTitle
 End Function
 '**********************************************************************