Option Explicit

Sub LockCells(SheetName As String, InputDate As Date)
	'-----------------------------------------------------------------------------------------------------------------------------'
	'Please send any questions or feedback to cmcguire@mvc.on.ca
	'-----------------------------------------------------------------------------------------------------------------------------'
	'The sheet from 7 days ago is locked and the weather cells on the current day's sheet are locked and a backup is saved.
	'-----------------------------------------------------------------------------------------------------------------------------'
	'The PrevDate variable is used to protect the 7 day data.
	Dim PrevDate As String
	'The PrevFound variable is used to determine whether or not the PrevDate sheet was found (and therefore exists)
	Dim PrevFound As Boolean
	'The 'DateAdd' function is used to calculate the date from 7 days ago.

	'The DateAdd function is nested inside the Format function and is used to calculate the previous date.
	'The Format function rearrages the date so that it can be processed by the KiWIS server.
	PrevDate = Format(DateAdd("d", -7, InputDate), "mmm d")
	PrevFound = False

	'This 'For' loop checks that the a sheet for the requested date does not already exist.
	Dim DPCsheet As Excel.Worksheet
	For Each DPCsheet In ThisWorkbook.Worksheets
		If DPCsheet.name = PrevDate Then
			PrevFound = True
		End If
	Next
	If Not PrevFound Then
		'If the sheet from 7 days ago does not exist, the user is alerted of the potential consequences
		Dim Answer As Integer
		Answer = MsgBox("A sheet for '" & PrevDate & "' was not found; as a result, this sheet's protection may not function as expected", vbOKOnly, "Missing Sheet for " & PrevDate)
	'End If

	'The With statement is used to ensure the macro does not modify other workbooks that may be open.
	Else
		With ThisWorkbook
			'The sheet from 7 days before the current date is unprotected.
			.Sheets(PrevDate).Unprotect
			'The previously unlocked cells are locked.
			.Sheets(PrevDate).Range("A1:M80").Locked = True
			'The entire sheet is now protected.
			.Sheets(PrevDate).Protect
		End With
	End If

	With ThisWorkbook
		'The cells containing gauge data are unlocked.
		.Sheets(SheetName).Range("A1:M80").Locked = False
		'The sheet is protected and by default, the weather data is locked.
		.Sheets(SheetName).Protect AllowInsertingRows:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
		
		'The daily planning cycle file is saved.
		.Save
		.Application.DisplayAlerts = False
		'A backup copy is saved to the Water Management Files folder and the local desktop.
		.SaveCopyAs "N:\common_folder\Water Management Files\Backup " & ThisWorkbook.name
		.SaveCopyAs CreateObject("WScript.Shell").SpecialFolders("Desktop") & Application.PathSeparator & "Backup " & ThisWorkbook.name
	End With
End Sub
