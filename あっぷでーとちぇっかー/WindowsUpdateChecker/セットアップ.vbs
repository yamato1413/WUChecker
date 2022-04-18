
'----------------------------------- 
' 2021�N7��30�� 14:01:59�쐬
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
		"C:\ISHII_Tools\�����ՂŁ[�Ƃ��������[\WindowsUpdateChecker\WindowsUpdateChecker.vbs" ,1,True
	wsh.Popup "�Z�b�g�A�b�v����"
End Sub

Sub Install()
	If Not fso.FolderExists("C:\ISHII_Tools") Then fso.CreateFolder "C:\ISHII_Tools"
	fso.CopyFolder wsh.CurrentDirectory,"C:\ISHII_Tools\�����ՂŁ[�Ƃ��������["
End Sub

Sub MakeTask(idx,time)
	Const TASK_TRIGGER_DAILY           = 2
	Const TASK_ACTION_EXEC             = 0
	Const TASK_CREATE_OR_UPDATE        = 6
	Const TASK_LOGON_INTERACTIVE_TOKEN = 3

	'// �^�X�N�X�P�W���[���ɐڑ�
	Set service = CreateObject("Schedule.Service")
	service.Connect

	'// �V�K�^�X�N�쐬
	Set rootFolder = service.GetFolder("\")
	Set taskDefinition = service.NewTask(0) 

	Set regInfo = taskDefinition.RegistrationInfo
	regInfo.Description = "���� " & time & " �ɍċN���̕K�v�����邩�ǂ������`�F�b�N���܂��B"
	regInfo.Author = "�Έ䑾�m"

	Set settings = taskDefinition.Settings
	settings.Enabled = True
	settings.StartWhenAvailable = True
	settings.Hidden = False

	'�g���K�[�̐ݒ�
	Set trigger = taskDefinition.Triggers.Create(TASK_TRIGGER_DAILY)
	trigger.StartBoundary = FormatTime(Date, time)
	trigger.DaysInterval = 1
	trigger.Id = "DailyTriggerId"
	trigger.Enabled = True
	
	'// ���s����t�@�C����ݒ�
	Set Action = taskDefinition.Actions.Create(TASK_ACTION_EXEC)
	Set fso = CreateObject("Scripting.FileSystemObject")
	Action.Path = "C:\ISHII_Tools\�����ՂŁ[�Ƃ��������[\WindowsUpdateChecker\WindowsUpdateChecker.vbs"
	
	'// �^�X�N��o�^
	rootFolder.RegisterTaskDefinition "Windows�A�b�v�f�[�g�`�F�b�N" & idx, taskDefinition, _ 
                                          TASK_CREATE_OR_UPDATE, , , _
                                          TASK_LOGON_INTERACTIVE_TOKEN
End Sub

Function FormatTime(date,time) 
'// �����̓��t�ƕ�����`���̎������󂯎���Đ��`���ꂽ�^�C���X�^���v�𕶎���`���ŕԂ��B
'// FormatTime(Date,"12:00:00") �� "2021-07-30T12:00:00"
 
	FormatTime = Year(date) & "-" & _
                         Right(0 & Month(date),2) & "-" & _
                         Right(0 & Day(date),2) & _
                         "T" & time
End Function