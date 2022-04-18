
'----------------------------------- 
' 2021年7月30日 17:08:00作成
'----------------------------------- 

Main
Sub Main()
On Error Resume Next
	Set service = CreateObject("Schedule.Service")
	service.Connect

	service.GetFolder("\").DeleteTask "Windowsアップデートチェック1",0
	service.GetFolder("\").DeleteTask "Windowsアップデートチェック2",0
	MsgBox "タスク削除完了"
End Sub
