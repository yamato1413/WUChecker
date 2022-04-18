
'----------------------------------- 
' 2021年7月30日 14:01:59作成
'----------------------------------- 

Const CheckTime1 = "11:50:00"
Const CheckTime2 = "15:00:00"

Dim wsh: Set wsh = CreateObject("WScript.Shell")
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")

Main
Sub Main()
	Install
	MakeTask 1,CheckTime1
	MakeTask 2,CheckTime2
	wsh.Run "C:\Windows\System32\cscript.exe " & _ 
		"C:\ISHII_Tools\あっぷでーとちぇっかー\WindowsUpdateChecker\WindowsUpdateChecker.vbs" ,1,True
	wsh.Popup "セットアップ完了"
End Sub

Sub Install()
	If Not fso.FolderExists("C:\ISHII_Tools") Then fso.CreateFolder "C:\ISHII_Tools"
	fso.CopyFolder wsh.CurrentDirectory,"C:\ISHII_Tools\あっぷでーとちぇっかー"
End Sub

Sub MakeTask(idx,time)
	Const TASK_TRIGGER_DAILY           = 2
	Const TASK_ACTION_EXEC             = 0
	Const TASK_CREATE_OR_UPDATE        = 6
	Const TASK_LOGON_INTERACTIVE_TOKEN = 3

	'// タスクスケジューラに接続
	Set service = CreateObject("Schedule.Service")
	service.Connect

	'// 新規タスク作成
	Set rootFolder = service.GetFolder("\")
	Set taskDefinition = service.NewTask(0) 

	Set regInfo = taskDefinition.RegistrationInfo
	regInfo.Description = "毎日 " & time & " に再起動の必要があるかどうかをチェックします。"
	regInfo.Author = "石井太知"

	Set settings = taskDefinition.Settings
	settings.Enabled = True
	settings.StartWhenAvailable = True
	settings.Hidden = False

	'トリガーの設定
	Set trigger = taskDefinition.Triggers.Create(TASK_TRIGGER_DAILY)
	trigger.StartBoundary = FormatTime(Date, time)
	trigger.DaysInterval = 1
	trigger.Id = "DailyTriggerId"
	trigger.Enabled = True
	
	'// 実行するファイルを設定
	Set Action = taskDefinition.Actions.Create(TASK_ACTION_EXEC)
	Set fso = CreateObject("Scripting.FileSystemObject")
	Action.Path = "C:\ISHII_Tools\あっぷでーとちぇっかー\WindowsUpdateChecker\WindowsUpdateChecker.vbs"
	
	'// タスクを登録
	rootFolder.RegisterTaskDefinition "Windowsアップデートチェック" & idx, taskDefinition, _ 
                                          TASK_CREATE_OR_UPDATE, , , _
                                          TASK_LOGON_INTERACTIVE_TOKEN
End Sub

Function FormatTime(date,time) 
'// 今日の日付と文字列形式の時刻を受け取って整形されたタイムスタンプを文字列形式で返す。
'// FormatTime(Date,"12:00:00") ⇒ "2021-07-30T12:00:00"
 
	FormatTime = Year(date) & "-" & _
                         Right(0 & Month(date),2) & "-" & _
                         Right(0 & Day(date),2) & _
                         "T" & time
End Function