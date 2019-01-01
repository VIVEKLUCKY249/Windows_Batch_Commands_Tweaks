length = InputBox("Length:")
' This function will generate random strong password of given length.
Function RndPassword(vLength)
	Const LETTERS = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789~-_+&*^$%#@!"
	For x=1 To vLength
		password = password & Mid(LETTERS, RandomNumber(1, Len(LETTERS)), 1)
	Next
	RndPassword = password
End Function
CopyToClipboard(RndPassword(length))
MsgBox "Generated Password is copied to clipboard."
' MsgBox "Generated Password:" & RndPassword(length)

Function RandomNumber(lowerLim, upperLim)
	Randomize
	RandomNumber = Int(((upperLim-lowerLim+1)* Rnd())+ lowerLim)
End Function

Sub CopyToClipboard(textToCopy)
	Set objHTML = CreateObject("InternetExplorer.Application")
	objHTML.Navigate ("about:blank")
	objHTML.Document.ParentWindow.ClipboardData.SetData "Text", textToCopy
	ClipboardText = objHTML.Document.ParentWindow.ClipboardData.GetData("Text")
	objHTML.Document.Close
	objHTML.Quit
	Set objHTML = Nothing
	CloseIEApplication
	CloseIEBrowser
End Sub

Function CloseIEApplication
	Set wmi = GetObject("winmgmts://./root/cimv2")
	browser = "iexplore.exe"
	qry = "SELECT * FROM Win32_Process WHERE Name='" & browser & "'"
	For Each p In wmi.ExecQuery(qry)
		p.Terminate
	Next
End Function

Sub CloseIEBrowser
	Dim WshShell, oExec
	Set WshShell = CreateObject("WScript.Shell")
	Set oExec = WshShell.Exec("taskkill /fi ""imagename eq iexplore.exe""")
	Do While oExec.Status = 0
		WScript.Sleep 100
	Loop
End Sub