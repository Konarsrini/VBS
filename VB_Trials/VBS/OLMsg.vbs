strTo  = "srinconnect@gmail.com"
strSub = "TestSub"
strMsg = "Sample Body"
SendMail strTo,strSub,strMsg
Private Function SendMail(strTo, strSub, strMsg)
    Dim olApp' As Outlook.Application
    Dim olNS' As NameSpace
    Dim olMessage' As MailItem
    
    Set olApp = CreateObject("Outlook.Application")
    Set olNS = olApp.GetNamespace("MAPI")
    
    Set olMessage = olApp.CreateItem(olMailItem)
    olMessage.To = strTo
    olMessage.Subject = strSub
    olMessage.Body = strMsg
    olMessage.Send
    
    olApp.Quit
    
    Set olApp = Nothing
    Set olNS = Nothing
    Set olMessage = Nothing
    
End Function