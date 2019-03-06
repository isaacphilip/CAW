Include("WebDriver.vbs")

Sub Include(sInstFile)
	Dim f, s, oFSO
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	On Error Resume Next
	If oFSO.FileExists(sInstFile) Then
		Set f = oFSO.OpenTextFile(sInstFile)
		s = f.ReadAll
		f.Close
		ExecuteGlobal s
	End If
	On Error Goto 0
	Set f = Nothing
	Set oFSO = Nothing
End Sub


Set Driver = New WebDriver	
    'Driver.connect "127.0.0.1","4444","firefoxproxy", ""
    Driver.navigateTo "http://www.google.com"	
	WScript.Sleep 4000
    MsgBox "Retrieve the URL of the current page: " & Driver.getCurrentUrl()
    MsgBox Driver.executeScript("alert('test')","")

'Set Element  = Driver.findElementBy(Driver.name,"q")
'    Element.sendKeys "VBScript"
'    Element.submit