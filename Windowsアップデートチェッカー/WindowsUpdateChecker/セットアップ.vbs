
'----------------------------------- 
' 2021�N7��30�� 14:01:59�쐬
'----------------------------------- 

Dim wsh: Set wsh = CreateObject("WScript.Shell")
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")

Main
Sub Main()
	Call Install
	Call SetupShortcut
	wsh.Run "C:\ISHII_Tools\Windows�A�b�v�f�[�g�`�F�b�J�[\WindowsUpdateChecker\WUChecker.exe" , 0 ,False
	wsh.Popup "�Z�b�g�A�b�v����"
End Sub

Sub Install()
	If Not fso.FolderExists("C:\ISHII_Tools") Then fso.CreateFolder "C:\ISHII_Tools"
	fso.CopyFolder wsh.CurrentDirectory,"C:\ISHII_Tools\Windows�A�b�v�f�[�g�`�F�b�J�["
End Sub

Sub SetupShortcut()
	Dim path_shortcut, path_tool
	path_shortcut = "C:\ISHII_Tools\Windows�A�b�v�f�[�g�`�F�b�J�[\WindowsUpdateChecker\WUChecker.lnk"
	path_tool = "C:\ISHII_Tools\Windows�A�b�v�f�[�g�`�F�b�J�[\WindowsUpdateChecker\WUChecker.exe"
	
	Dim shortcut
	Set shortcut = wsh.CreateShortcut(path_shortcut)
	shortcut.TargetPath = path_tool
	shortcut.Save
	
	fso.CopyFile  path_shortcut, wsh.SpecialFolders("Startup") & "\WUChecker.lnk", True
End Sub