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

	'This 'For' loop checks that the sheet to be locked does in fact exist
	Dim DPCsheet As Excel.Worksheet
	For Each DPCsheet In ThisWorkbook.Worksheets
		If DPCsheet.name = PrevDate Then
			PrevFound = True
			Exit For 'If we've already found PrevDate, we don't need to continue looking for it
		End If
	Next
	If Not PrevFound Then
		'If the sheet from 7 days ago does not exist, the user is alerted of the potential consequences
		Dim Answer As Integer
		If Not IsAuto Then _
			Answer = MsgBox("A sheet for '" & PrevDate & "' was not found; as a result, '" & PrevDate & "' cannot be properly locked", vbOKOnly, "Missing Sheet for " & PrevDate)

	'The With statement is used to ensure the macro does not modify other workbooks that may be open.
	Else
		With ThisWorkbook
			'The sheet from 7 days before the current date is unprotected.
			.Sheets(PrevDate).Unprotect
			'The previously unlocked cells are locked.
			.Sheets(PrevDate).Range("A1:M" & AccuStart - 1).Locked = True
			'The entire sheet is now protected.
			.Sheets(PrevDate).Protect
		End With
	End If

	With ThisWorkbook
		'The cells containing gauge data are unlocked.
		.Sheets(SheetName).Range("A1:M" & AccuStart - 1).Locked = False
		'The sheet is protected and by default, the weather data is locked.
		.Sheets(SheetName).Protect AllowInsertingRows:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
		
		'The Daily Planning Cycle file is saved.
		On Error Resume Next
		.Save
		If Err.Number <> 0 And Not IsAuto Then _
			MsgBox "Failed to save because no network was found"
		On Error GoTo 0
		.Application.DisplayAlerts = False
		'A backup copy is saved to the 'Water Management Files' folder and the local desktop.
		On Error Resume Next
		.SaveCopyAs "N:\common_folder\Water Management Files\Backup " & ThisWorkbook.name
		If Err.Number <> 0 And Not IsAuto Then _
			MsgBox "Failed to save because no network was found"
		On Error GoTo 0
		.SaveCopyAs "C:\Users\Public\Documents\Backup " & ThisWorkbook.name
		.SaveCopyAs CreateObject("WScript.Shell").SpecialFolders("Desktop") & Application.PathSeparator & "Backup " & ThisWorkbook.name
	End With
End Sub