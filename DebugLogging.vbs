'/* 
' * DebugMsg and PrintMsg are used in conjunction in order to log messages to be returned when the current macro is finished.
' * DebugMsg stores the string data and PrintMsg modifies and returns it.  When PrintMsg is given no input, it will only 
' * return.
' * 
' * @param Txt: String containing the next message to be logged.  Defaults to empty.
' * 
' * @return the current value of DebugMsg
' */
Private DebugMsg As String
Private output As Object
Public Function PrintMsg(Optional Txt As String = "", Optional NoStatus As Boolean = False)
	If Not output Is Nothing And Txt <> "" Then _
		output.WriteLine "[" & Now & "] " & Txt
	If Not Txt = "" Then _
		DebugMsg = DebugMsg & "[" & Now & "] " & Txt & vbCrLf 'Don't add text if no input is given
	PrintMsg = DebugMsg 'Return to get the value of DebugMsg
	If Txt = "" Or NoStatus Then _
		Exit Function
	Call ChangeStatus(Txt)
End Function


Public Function SetOutput(Optional obj As Object = Nothing)
	Set output = obj
End Function

Public Sub Clear()
	DebugMsg = ""
End Sub

Public Sub Erred()
	If Err.Number = 0 Then _
		Exit Sub
	Dim ErrMsg As String

	ErrMsg = "Error " & Err.Number & " at line " & Erl & ": " & Err.Description
	If Erl = 0 Then _
		ErrMsg = "Error " & Err.Number & ": " & Err.Description
	PrintMsg(ErrMsg)
End Sub