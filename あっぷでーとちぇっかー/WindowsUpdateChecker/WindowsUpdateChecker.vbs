
'----------------------------------- 
' 2021�N7��29�� 11:55:35�쐬
'----------------------------------- 

Dim wsh: Set wsh = CreateObject("WScript.Shell")

Main
Sub Main()
	'�����̃��W�X�g���͍ċN���̕K�v������Ƃ��ɂ̂ݑ��݂��܂��B
	Const REG_RebootPending = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending\"
	Const REG_RebootRequired = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired\"
	
	If RegExist(REG_RebootPending) Or RegExist(REG_RebootRequired) Then
		wsh.Popup "�V���b�g�_�E�����ɍX�V������\��������܂��B" & vbCrLf & _
			"�x�e���ԓ��ɍċN�����������߂��܂��B",, _
			"�����ՂŁ[�Ƃ��������[", _
			vbExclamation + vbSystemModal
	Else
		If InStr(LCase(WScript.FullName),"cscript") Then
			WScript.Echo "�V���b�g�_�E������Windows�A�b�v�f�[�g�Ȃ�" & vbCrLf
			WScript.Sleep 2000
		End IF
	End If

	Set wsh = Nothing
End Sub

Function RegExist(RegPath)
	On Error Resume Next
	wsh.RegRead RegPath
	If Err Then RegExist = False Else RegExist = True
	Err.Clear
End Function