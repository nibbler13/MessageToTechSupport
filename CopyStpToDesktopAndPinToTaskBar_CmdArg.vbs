Option Explicit
On Error Resume Next

Set args = Wscript.Arguments
Dim path

If Not Wscript.Arguments.Item(0) = "" Then
	path = Wscript.Arguments.Item(0)
Else
	Wscript.Quit
End If

Dim ShellApp, FSO, Desktop, objShell, strPath, LnkFile
Set ShellApp = CreateObject("Shell.Application")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set objShell = Wscript.CreateObject("Wscript.Shell")
strPath = objShell.SpecialFolders("Desktop")

'"\\nnkk-fs\share\��������� � ������ ����������� ���������.exe"
If FSO.FileExists(path) then
	FSO.CopyFile path, strPath & "\"
Else
	Wscript.Quit
End If

Set Desktop =  ShellApp.NameSpace(strPath)
LnkFile = strPath & "\��������� � ������ ����������� ���������.exe"

If(FSO.FileExists(LnkFile)) Then
    Dim tmp, verb, desktopImtes, item
    Set desktopImtes = Desktop.Items()

    For Each item in desktopImtes
        If (item.Name = "��������� � ������ ����������� ���������.exe") _
	Or (item.Name = "��������� � ������ ����������� ���������") Then
            For Each verb in item.Verbs
                If (verb.Name = "��������� �� &������ �����") Then
                    verb.DoIt
                End If
            Next
        End If
    Next
End If

Set FSO = Nothing
Set ShellApp = Nothing