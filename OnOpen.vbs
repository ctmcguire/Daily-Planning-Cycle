Public TimeOut As Boolean

Public Sub ShutdownWarning()
	ShutdownWarningForm.show
End Sub

Public Sub ShutdownAction()
	If TimeOut Then _
		Unload ShutdownWarningForm
	ThisWorkbook.Saved = True 'NOTE: this makes Excel THINK the file is saved, so that the "save your work" popup doesn't appear.  IT DOESN'T ACTUALLY SAVE ANYTHING.
	ThisWorkbook.Close
End Sub

Private Sub TestShutdown()
	Application.OnTime Now + TimeValue("00:00:30"), "ShutdownWarning"
End Sub

Private Sub Auto_Open()
	Application.OnTime "7:20:00", "ShutdownWarning"
End Sub