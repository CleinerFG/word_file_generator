Set objWord = CreateObject("Word.Application")
objWord.Visible = False

Set objFSO = CreateObject("Scripting.FileSystemObject")
strFolder = ".\arq_modificados" ' Diretório que você deseja listar e imprimir

For Each objFile In objFSO.GetFolder(strFolder).Files
    If LCase(objFSO.GetExtensionName(objFile)) = "docx" Then
        objWord.Documents.Open objFile.Path
        objWord.PrintOut
        objWord.ActiveDocument.Close False
    End If
Next

objWord.Quit
