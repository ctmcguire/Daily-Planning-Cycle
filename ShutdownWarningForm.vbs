Private CloseTime As Date

Private Sub UserForm_Initialize()
	TimeOut = True
	CloseTime = Now + TimeValue("00:03:00")
	Application.OnTime CloseTime, "ShutdownAction"
End Sub

Private Sub OkayButton_Click()
	Unload Me
End Sub

Private Sub UserForm_Terminate()
	TimeOut = False
End Sub