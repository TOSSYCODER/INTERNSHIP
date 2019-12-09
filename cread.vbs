
Dim Textstring

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objText = objFSO.OpenTextFile("File.txt", 1)

Textstring = objText.ReadAll