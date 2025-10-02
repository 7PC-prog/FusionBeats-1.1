Option Explicit
Dim objShell, objFSO, strUserDocs, strProgPath, strDesktop
Dim objFolderSource, objFile, objFolder, objShortcut, strTargetPath, strLinkPath, strIconPath

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' ���� ��������� � "���������"
strUserDocs = objShell.SpecialFolders("MyDocuments")
strProgPath = strUserDocs & "\Floppotron_1.0"

' �������� � ���� ����� ��� ����, �������
If objFSO.FolderExists(strProgPath) Then
    MsgBox "��������� ��� ����������� �: " & strProgPath, vbExclamation, "Floppotron Setup"
    WScript.Quit
End If

' ������� ����� ���������
objFSO.CreateFolder strProgPath

' �������� ��� ����� �� ������� ����� (����� installer.vbs)
Set objFolderSource = objFSO.GetFolder(".")
For Each objFile In objFolderSource.Files
    If LCase(objFSO.GetFileName(objFile)) <> "installer.vbs" Then
        objFSO.CopyFile objFile.Path, strProgPath & "\", True
    End If
Next

' �������� ����� Floppotron_Folder, ���� ��� ����������
If objFSO.FolderExists(".\Floppotron_Folder") Then
    CopyFolderRecursive ".\FusionBeats_Folder", strProgPath & "\FusionBeats_Folder"
End If

' ������� ����� �� ������� �����
strDesktop = objShell.SpecialFolders("Desktop")
strTargetPath = strProgPath & "\FloppotronGUI.hta"  ' ��������� GUI
strLinkPath = strDesktop & "\FusionBeats 1.1 Release.lnk"
strIconPath = strProgPath & "\Icon.bmp"  ' ������ � ����� ���������

' ���������, ���� ������ ����
If Not objFSO.FileExists(strIconPath) Then
    strIconPath = "" ' ���� ���, �� ����� �����������
End If

Set objShortcut = objShell.CreateShortcut(strLinkPath)
objShortcut.TargetPath = strTargetPath
objShortcut.WorkingDirectory = strProgPath
If strIconPath <> "" Then objShortcut.IconLocation = strIconPath
objShortcut.Save

MsgBox "��������� ���������!" & vbCrLf & "����: " & strProgPath, vbInformation, "Floppotron Setup"

'---------------------------
' ������� ������������ ����������� �����
Sub CopyFolderRecursive(ByVal sourceFolder, ByVal targetFolder)
    Dim folder, file
    objFSO.CreateFolder targetFolder
    Set folder = objFSO.GetFolder(sourceFolder)
    
    ' �������� �����
    For Each file In folder.Files
        objFSO.CopyFile file.Path, targetFolder & "\", True
    Next
    
    ' �������� ��������� ����� ����������
    For Each folder In folder.SubFolders
        CopyFolderRecursive folder.Path, targetFolder & "\" & objFSO.GetFileName(folder.Path)
    Next
End Sub
