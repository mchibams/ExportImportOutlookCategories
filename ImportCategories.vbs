On Error Resume Next
Const IMPORT_FILE = "C:\temp\categories.txt"
Dim olkApp 'As Outlook.Application
Dim objFSO 'As FileSystemObject
Dim stmFile 'As TextStream
Dim strLine 'As String
Dim arrField 'As Variant
Dim objCategory 'As Outlook.Category
'
Set olkApp = CreateObject("Outlook.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set stmFile = objFSO.OpenTextFile(IMPORT_FILE, 1 ) ' ForReading
While Not stmFile.AtEndOfStream
    strLine = stmFile.ReadLine
    arrField = Split(strLine, ";")
    olkApp.Session.Categories.Add arrField(2), arrField(0), arrField(1)
    If Err.Number <> 0 Then ' If the category already exists, overwrite it.
        Set objCategory = olkApp.Session.Categories.Item(arrField(2))
        objCategory.Color = arrField(0)
        objCategory.ShortcutKey = arrField(1)
    End If
Wend
stmFile.Close
