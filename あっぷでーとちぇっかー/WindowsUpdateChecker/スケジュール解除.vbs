
'----------------------------------- 
' 2021�N7��30�� 17:08:00�쐬
'----------------------------------- 

Main
Sub Main()
On Error Resume Next
	Set service = CreateObject("Schedule.Service")
	service.Connect

	service.GetFolder("\").DeleteTask "Windows�A�b�v�f�[�g�`�F�b�N1",0
	service.GetFolder("\").DeleteTask "Windows�A�b�v�f�[�g�`�F�b�N2",0
	MsgBox "�^�X�N�폜����"
End Sub
