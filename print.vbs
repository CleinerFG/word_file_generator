' Print all word files

Set objWord = CreateObject("Word.Application")
objWord.Visible = False

Set objFSO = CreateObject("Scripting.FileSystemObject")
strFolder = ".\output"

' Set duplex config
Set objPrintOptions = objWord.ActivePrinter
objPrintOptions.Duplex = 2 ' 1 for single-sided, 2 for double-sided.

For Each objFile In objFSO.GetFolder(strFolder).Files
    If LCase(objFSO.GetExtensionName(objFile)) = "docx" Then
        objWord.Documents.Open objFile.Path
        objWord.PrintOut
        objWord.ActiveDocument.Close False
    End If
Next

objWord.Quit
