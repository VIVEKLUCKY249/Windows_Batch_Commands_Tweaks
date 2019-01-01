length = InputBox("Length:")

Function MakeRandomeDigits(numLength)
	ReDim storedDigits(-1)
	Dim numberString
	Randomize
	For i = 1 to numLength
		ReDim Preserve storedDigits(UBound(storedDigits) + 1)
		storedDigits(UBound(storedDigits)) = Int(9 * Rnd()) + 1
	Next
	For j = 0 to UBound(storedDigits)
		numberString = numberString & storedDigits(j)
	Next
	MakeRandomeDigits = numberString
End Function
CopyToClipboard(MakeRandomeDigits(length))
MsgBox "Generated Passcode is copied to clipboard."

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