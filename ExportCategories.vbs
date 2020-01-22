Const EXPORT_FILE = "C:\temp\categories.txt"
Dim olkApp 'As Outlook.Application
Dim objFSO 'As FileSystemObject
Dim stmFile 'As TextStream
Dim objCategory 'As Outlook.Category
'
Set olkApp = CreateObject("Outlook.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set stmFile = objFSO.CreateTextFile(EXPORT_FILE, True, False)
For Each objCategory In olkApp.Session.Categories
    stmFile.WriteLine objCategory.Color & ";" & objCategory.ShortcutKey & ";" & objCategory.Name
Next
stmFile.Close
