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
Public Function PrintMsg(Optional Txt As String = "")
	If Not Txt = "" Then
		DebugMsg = DebugMsg & "[" & Now & "] " & Txt & vbCrLf 'Don't add text if no input is given
	PrintMsg = DebugMsg 'Return to get the value of DebugMsg
End Function

Public Sub Erred()
	If Err.Number = 0 Then _
		Exit Sub
	Dim ErrMsg As String

	ErrMsg = "Error " & Err.Number & " at line " & Erl & ": " & Err.Description
	If Erl = 0 Then _
		ErrMsg = "Error " & Err.Number & ": " & Err.Description
	PrintMsg(ErrMsg)
End Sub