Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = ".\backgrounds"

copyImageTo = ".\background.jpg"

Set objFolder = objFSO.GetFolder(objStartFolder)
Set colFiles = objFolder.Files

intHighNumber = (colFiles.Count - 1)
intLowNumber = 0

Randomize
randomNumber = Int((intHighNumber - intLowNumber + 1) * Rnd + intLowNumber)
Dim i
i = 0

For Each image in colFiles
    If i = randomNumber Then
        Wscript.Echo "files: " & image.name
        image.Copy(copyImageTo)
        Exit For
    End If
    i = i + 1
Next