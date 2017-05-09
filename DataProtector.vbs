Option Explicit

Sub LockCells(SheetName As String, InputDate As Date)
	'-----------------------------------------------------------------------------------------------------------------------------'
	'Please send any questions or feedback to cmcguire@mvc.on.ca
	'-----------------------------------------------------------------------------------------------------------------------------'
	'The sheet from 7 days ago is locked and the weather cells on the current day's sheet are locked and a backup is saved.
	'-----------------------------------------------------------------------------------------------------------------------------'
	'The PrevDate variable is used to protect the 7 day data.
	Dim PrevDate As String
	'The 'DateAdd' function is used to calculate the date from 7 days ago.

	'The DateAdd function is nested inside the Format function and is used to calculate the previous date.
	'The Format function rearrages the date so that it can be processed by the KiWIS server.
	PrevDate = Format(DateAdd("d", -7, InputDate), "mmm d")

	'The With statement is used to ensure the macro does not modify other workbooks that may be open.
	With ThisWorkbook
		'The sheet from 7 days before the current date is unprotected.
		.Sheets(PrevDate).Unprotect
		'The previously unlocked cells are locked.
		.Sheets(PrevDate).Range("A1:M80").Locked = True
		'The entire sheet is now protected.
		.Sheets(PrevDate).Protect
		
		'The cells containing gauge data are unlocked.
		.Sheets(SheetName).Range("A1:M80").Locked = False
		'The sheet is protected and by default, the weather data is locked.
		.Sheets(SheetName).Protect AllowInsertingRows:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
		
		'The daily planning cycle file is saved.
		.Save
		.Application.DisplayAlerts = False
		'A backup copy is saved to the Water Management Files folder and the local desktop.
		.SaveCopyAs "N:\common_folder\Water Management Files\Backup Automated Daily Planning Cycle  - 2016.xlsm"
		.SaveCopyAs CreateObject("WScript.Shell").SpecialFolders("Desktop") & Application.PathSeparator & "Backup Automated Daily Planning Cycle  - 2016.xlsm"
	End With
End Sub
