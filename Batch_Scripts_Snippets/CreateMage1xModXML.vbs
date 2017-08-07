' VBTab 	A Tab character [Chr(9)]
' VBCr 	A carriage return [Chr(13)]
' VBCrLf 	A carriage return and line feed [Chr(13) + Chr(10)]
' vbBack 	A backspace character [Chr(8)]
' vbLf 	A linefeed [Chr(10)]
' vbNewLine 	A platform-specific new line character, either [Chr(13) + Chr(10)] or [Chr(13)]
' vbNullChar 	A null character of value 0 [Chr(0)]
' vbNullString 	A string of value 0 [no Chr code]; note that this is not the same as “”

scriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
currFolder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
Set xmlDoc = CreateObject("Msxml2.DOMDocument.3.0")

packageName = InputBox("Package Name:")
moduleName = InputBox("Module Name:")
codePool = InputBox("CodePool:")
moduleVersion = InputBox("Module version:")

If packageName = "" Or moduleName = "" Or codePool = "" Then
	WScript.Quit
End If

If moduleVersion = "" Then
	moduleVersion = "1.0.0"
End If

xmlString = "<?xml version=" & chr(34) & "1.0" & chr(34) & "?>" & vbCr & _
"<config>" & vbCr & _
VBTab & "<modules>" & vbCr & _
VBTab & VBTab & "<" & packageName & "_" & moduleName & ">" & vbCr & _
VBTab & VBTab & VBTab & "<active>true</active>" & vbCr & _
VBTab & VBTab & VBTab & "<codePool>" & codePool & "</codePool>" & vbCr & _
VBTab & VBTab & VBTab & "<version>" & moduleVersion & "</version>" & vbCr & _
VBTab & VBTab & "</" & packageName & "_" & moduleName & ">" & vbCr & _
VBTab & "</modules>" & vbCr & _
"</config>"
'xmlDoc.appendChild(xmlString)
xmlFilePath = scriptDir & "\" & packageName & "_" & moduleName & ".xml"
xmlDoc.Save xmlFilePath
Set fS = CreateObject("Scripting.FileSystemObject")
Set fSO = fS.OpenTextFile(xmlFilePath, 2, false)
fSO.Write xmlString
fSO.Close()