
Option Explicit

Private Sub CommandButton1_Click()

	InputHour = HourPicker.DTPicker5.Value
	InputTime = Format(HourPicker.DTPicker5.Value, "ham/pm mmm d")

	'The form is unloaded to free up memory.
	Unload Me
End Sub

'This Private Subroutine cancels out of the DPCUpdate when the close button is pressed
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
	InputTime = "cancel"
	Unload Me
End Sub

Private Sub UserForm_Initialize()
	HourPicker.DTPicker5.Value = Date + TimeSerial(Hour(Now), 0, 0)
End Sub
