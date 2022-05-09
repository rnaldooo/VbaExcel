
''(Include a reference to Microsoft XML v3, I saved your xml to a file on my desktop)

Dim xmlDoc As DOMDocument30
Set xmlDoc = New DOMDocument30
xmlDoc.Load ("C:\users\jon\desktop\test.xml")

Dim id As String
id = xmlDoc.SelectSingleNode("//GetUserInfo/User").Attributes.getNamedItem("ID").Text