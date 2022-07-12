Dim IE
 Set IE = CreateObject("InternetExplorer.Application")
 IE.Visible = 1 
 IE.navigate "http://SolidWorks.com"
 Do While (IE.Busy)
   WScript.Sleep 10
 Loop
msgbox "ready"
 'Set Helem = IE.document.getElementByID("formUsername")
 'Helem.Value = "username" ' change this to yours
 'Set Helem = IE.document.getElementByID("formPassword")
 'Helem.Value = "password" ' change this to yours
 'Set Helem = IE.document.Forms(0)
 'Helem.Submit