Option Explicit
Dim objShell, objFSO, strUserDocs, strProgPath, strDesktop
Dim objFolderSource, objFile, objFolder, objShortcut, strTargetPath, strLinkPath, strIconPath

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Путь установки в "Документы"
strUserDocs = objShell.SpecialFolders("MyDocuments")
strProgPath = strUserDocs & "\Floppotron_1.0"

' Проверка — если папка уже есть, выходим
If objFSO.FolderExists(strProgPath) Then
    MsgBox "Программа уже установлена в: " & strProgPath, vbExclamation, "Floppotron Setup"
    WScript.Quit
End If

' Создаем папку установки
objFSO.CreateFolder strProgPath

' Копируем все файлы из текущей папки (кроме installer.vbs)
Set objFolderSource = objFSO.GetFolder(".")
For Each objFile In objFolderSource.Files
    If LCase(objFSO.GetFileName(objFile)) <> "installer.vbs" Then
        objFSO.CopyFile objFile.Path, strProgPath & "\", True
    End If
Next

' Копируем папку Floppotron_Folder, если она существует
If objFSO.FolderExists(".\Floppotron_Folder") Then
    CopyFolderRecursive ".\FusionBeats_Folder", strProgPath & "\FusionBeats_Folder"
End If

' Создаем ярлык на рабочем столе
strDesktop = objShell.SpecialFolders("Desktop")
strTargetPath = strProgPath & "\FloppotronGUI.hta"  ' запускаем GUI
strLinkPath = strDesktop & "\FusionBeats 1.1 Release.lnk"
strIconPath = strProgPath & "\Icon.bmp"  ' иконка в папке установки

' Проверяем, если иконка есть
If Not objFSO.FileExists(strIconPath) Then
    strIconPath = "" ' если нет, то будет стандартная
End If

Set objShortcut = objShell.CreateShortcut(strLinkPath)
objShortcut.TargetPath = strTargetPath
objShortcut.WorkingDirectory = strProgPath
If strIconPath <> "" Then objShortcut.IconLocation = strIconPath
objShortcut.Save

MsgBox "Установка завершена!" & vbCrLf & "Путь: " & strProgPath, vbInformation, "Floppotron Setup"

'---------------------------
' Функция рекурсивного копирования папки
Sub CopyFolderRecursive(ByVal sourceFolder, ByVal targetFolder)
    Dim folder, file
    objFSO.CreateFolder targetFolder
    Set folder = objFSO.GetFolder(sourceFolder)
    
    ' Копируем файлы
    For Each file In folder.Files
        objFSO.CopyFile file.Path, targetFolder & "\", True
    Next
    
    ' Копируем вложенные папки рекурсивно
    For Each folder In folder.SubFolders
        CopyFolderRecursive folder.Path, targetFolder & "\" & objFSO.GetFileName(folder.Path)
    Next
End Sub
