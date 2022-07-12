Set objOutlook = CreateObject("Outlook.Application")
Set ObjB = objOutlook.Session.Accounts
    For I = 1 To ObjB.Count
        MsgBox ObjB.Item(I)
    Next

'Set myNameSpace = objOutlook.GetNameSpace("MAPI")
'Set Folders = myNameSpace.GetDefaultFolder(3)
'intMainFoldCount = Folders.ShowItemCount
'msgbox intMainFoldCount


'strFolderName = Folders.Parent
msgbox strFolderName

For k = 1 to intMainFoldCount
	Set objMailbox = myNameSpace.Folders(Folders)
	Set objFolder = objMailbox.Folders("Inbox")
	Set objFolder1 = objFolder.Folders
	For Each Folder in objFolder1
		Set myItems = objFolder1.Items
		If strB = "" Then
			strB = Folder.Name
		Else
			strB = strB&";"&Folder.Name
		End If
	Next
Next	  
'msgbox strB
