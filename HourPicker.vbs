
Option Explicit

Private Sub CommandButton1_Click()

	SheetName = Format(HourPicker.DTPicker3.Value, "ham/pm mmm d")
	SheetDay = HourPicker.DTPicker3.Value

	'The form is unloaded to free up memory.
	Unload Me
End Sub

Private Sub UserForm_Initialize()
	HourPicker.DTPicker3.Value = Date + TimeSerial(Hour(Now), 0, 0)
End Sub