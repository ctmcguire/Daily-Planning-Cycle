Option Explicit

Sub LockCells(SheetName As String, InputDate As Date, Optional IsAuto As Boolean = False)
	'-----------------------------------------------------------------------------------------------------------------------------'
	'Please send any questions or feedback to cmcguire@mvc.on.ca
	'-----------------------------------------------------------------------------------------------------------------------------'
	'The sheet from 7 days ago is locked and the weather cells on the current day's sheet are locked and a backup is saved.
	'-----------------------------------------------------------------------------------------------------------------------------'
	'The PrevDate variable is used to protect the 7 day data.
	Dim PrevDate As String
	'The PrevFound variable is used to determine whether or not the PrevDate sheet was found (and therefore exists).
	Dim PrevFound As Boolean
	
	'The DateAdd function is nested inside the Format function and is used to calculate the previous date (7 days).
	'The Format function rearrages the date so that it matches the format of the sheet names.
	PrevDate = Format(DateAdd("d", -7, InputDate), "mmm d")
	PrevFound = False

	Call DebugLogging.PrintMsg("Checking if last week's sheet exists...")

	'This 'For' loop checks that the sheet to be locked does in fact exist
	Dim DPCsheet As Excel.Worksheet
	For Each DPCsheet In ThisWorkbook.Worksheets
		If DPCsheet.name = PrevDate Then
			PrevFound = True
			Exit For 'If we've already found PrevDate, we don't need to continue looking for it
		End If
	Next

	If Not PrevFound Then
		If Not IsAuto Then _
			MsgBox "A sheet for '" & PrevDate & "' was not found; as a result, '" & PrevDate & "' cannot be properly locked", vbOKOnly, "Missing Sheet for " & PrevDate
		Call DebugLogging.PrintMsg("Unable to find last week's sheet; missing sheet will not be locked.")
	Else
		Call DebugLogging.PrintMsg("Sheet found.  Locking...")

		'The With statement is used to ensure the macro does not modify other workbooks that may be open.
		With ThisWorkbook
			'The sheet from 7 days before the current date is unprotected.
			.Sheets(PrevDate).Unprotect
			'The previously unlocked cells are locked.
			.Sheets(PrevDate).Range(LockRange).Locked = True
			'The entire sheet is now protected.
			.Sheets(PrevDate).Protect
		End With
		Call DebugLogging.PrintMsg("Sheet successfully locked.")
	End If

	With ThisWorkbook
		Call DebugLogging.PrintMsg("Protecting worksheet...")

		.Sheets(SheetName).Range(LockRange).Locked = False 'The cells containing gauge data are unlocked.
		'The sheet is protected and by default, the weather data is locked.
		.Sheets(SheetName).Protect AllowInsertingRows:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True

		Call DebugLogging.PrintMsg("Worksheet protected.  Saving worksheet")

		On Error Resume Next
		.Save 'The Daily Planning Cycle file is saved.
		If Err.Number <> 0 Then
			If Not IsAuto Then _
				MsgBox "Failed to save because no network was found"
			DebugLogging.PrintMsg("Failed to save because no network was found.")
		End If
		On Error GoTo 0
		.Application.DisplayAlerts = False

		Call DebugLogging.PrintMsg("Saving server backup...")

		'A backup copy is saved to the 'Water Management Files' folder and the local desktop.
		On Error Resume Next
		.SaveCopyAs BackupFolder & "\Backup " & ThisWorkbook.name
		If Err.Number <> 0 Then
			If Not IsAuto Then _
				MsgBox "Failed to save because no network was found"
			DebugLogging.PrintMsg("Failed to save because no network was found")
		End If
		On Error GoTo 0

		Call DebugLogging.PrintMsg("Saving local backups...")

		.SaveCopyAs "C:\Users\Public\Documents\Backup " & ThisWorkbook.name
		.SaveCopyAs CreateObject("WScript.Shell").SpecialFolders("Desktop") & Application.PathSeparator & "Backup " & ThisWorkbook.name
	End With
End Sub

Private Function GetMnth(ShNm As String)
	GetMnth = Mid(ShNm,1,3)
	Dim i
	For Each i in Array("am ", "pm ")
		If 1 < InStr(2, Mid(ShNm,1,5), "" & i) And InStr(2, Mid(ShNm,1,5), "" & i) < 4 Then
			GetMnth = Mid(ShNm, InStr(2, Mid(ShNm,1,5), "" & i) + Len("" & i), 3)
			Exit Function
		End If
	Next
End Function
Function SaveNextYear()
	Dim FileName As String
	Call DebugLogging.Clear()
	With ThisWorkbook
		FileName = Replace(.name, Year(Now), Year(DateAdd("yyyy", 1, Now)))
		.SaveAs .path & "\" & FileName

		Dim DPCsheet As Excel.Worksheet
		For Each DPCsheet in .Worksheets
			Dim mnth
			With DPCsheet
				Dim ShNm As String
				ShNm = GetMnth(.name)
				For Each mnth In Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","")
					If ShNm = "" & mnth Then _
						Exit For
				Next
				If mnth = "" Then _
					Goto continue
				Application.DisplayAlerts = False
				DPCsheet.Delete
				Application.DisplayAlerts = True
				continue:
			End With
		Next
		.Save
	End With
	SaveNextYear = DebugLogging.PrintMsg()
End Function

Sub SavePreBackup(Optional IsAuto As Boolean = False)
	With ThisWorkbook
		Call DebugLogging.PrintMsg("Saving server pre-run backup...")
		.SaveCopyAs "\\APP-SERVER\Data_drive\common_folder\Water Management Files\Pre-Run Backup " & ThisWorkbook.name
		If Err.Number <> 0 Then
			If Not IsAuto Then _
				MsgBox "Failed to save because no network was found"
			DebugLogging.PrintMsg("Failed to save because no network was found")
		End If
		On Error GoTo 0

		Call DebugLogging.PrintMsg("Saving local pre-run backups...")

		.SaveCopyAs "C:\Users\Public\Documents\Pre-Run Backup " & ThisWorkbook.name
		.SaveCopyAs CreateObject("WScript.Shell").SpecialFolders("Desktop") & Application.PathSeparator & "Pre-Run Backup " & ThisWorkbook.name
	End With
End Sub

Sub EditCells(SheetName As String, Optional IsAuto As Boolean = False)
	Call DebugLogging.PrintMsg("Duplicate found.  Filling blanks" & Switch(IsAuto, "...", True, " as requested by user..."))
	With ThisWorkbook
		If Not IsAuto Then _
			Call SavePreBackup()
		.Sheets(SheetName).Unprotect
	End With
End Sub