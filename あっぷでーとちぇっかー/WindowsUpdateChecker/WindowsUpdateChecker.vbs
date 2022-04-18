
'----------------------------------- 
' 2021年7月29日 11:55:35作成
'----------------------------------- 

Dim wsh: Set wsh = CreateObject("WScript.Shell")

Main
Sub Main()
	'これらのレジストリは再起動の必要があるときにのみ存在します。
	Const REG_RebootPending = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending\"
	Const REG_RebootRequired = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired\"
	
	If RegExist(REG_RebootPending) Or RegExist(REG_RebootRequired) Then
		wsh.Popup "シャットダウン時に更新が入る可能性があります。" & vbCrLf & _
			"休憩時間等に再起動をおすすめします。",, _
			"あっぷでーとちぇっかー", _
			vbExclamation + vbSystemModal
	Else
		If InStr(LCase(WScript.FullName),"cscript") Then
			WScript.Echo "シャットダウン時のWindowsアップデートなし" & vbCrLf
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