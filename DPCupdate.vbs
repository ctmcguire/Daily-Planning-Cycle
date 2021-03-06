Option Explicit


'NOTE: Most of the Constant values used by UpdateDPC are created in the Gauges Macro

Public NextWeather As Integer

Public Const AccuCount As Integer = 5
Public Const TWNCount As Integer = 15
Public Const ECCount As Integer = 19

Public Const AccuStart As Integer = weeklyStart + weeklyCount + 12
Public Const TWNStart As Integer = AccuStart + AccuCount + 2
Public Const ECStart As Integer = TWNStart + TWNCount + 2
Public Const CloyneAccuStart As Integer = ECStart + ECCount + 2
Public Const CloyneTWNStart As Integer = CloyneAccuStart + AccuCount + 2

'The date picker assigns a value to the public variable 'SheetName'.
'The variable is public so that it can be used by all the called functions.
Public SheetName As String
Public SheetDay As Date

Private Sub Start()
	'The Status Bar is located on the bottom left corner of the Excel window.  It's default status is 'READY'.
	Call ChangeStatus("Processing Request...") 'The Status Bar Displays 'Processing Request...' until the UpdateDPC subroutine has ended.
	Application.ScreenUpdating = False 'Screen Updating is turned off to speed up the processing time.
	Application.Calculation = xlCalculationManual 'Sheet calculations are turned off to speed up the processing time.
	Call DebugLogging.Clear 'Clear the debug log (in case it isn't empty already)
End Sub

Private Sub Finish()
	'The previously adjusted modes are returned to their default state.
	Call ChangeStatus
	Application.Calculation = xlCalculationAutomatic
	Application.ScreenUpdating = True
End Sub

Sub UpdateDPCByDate()
	Call Start
	SheetName = "cancel" 'If no date is chosen, this variable is assigned to exit the macro.
	'-------------------------------------------------------------------------------------------------------------------------------------------------'
	'Instructions to install the Date Picker control can be found by right clicking on the 'DatePicker' form and selecting 'View Code'.
	DatePicker.Show 'The Date Picker form is shown and the user inputs a date.

	Call UpdateDPC(SheetName, SheetDay)
	Call Finish
End Sub

Sub UpdateDPCByHour()
	Call Start
	SheetName = "cancel" 'If no time is chosen, this variable is assigned to exit the macro.
	'-------------------------------------------------------------------------------------------------------------------------------------------------'
	'Instructions to install the Date Picker control can be found by right clicking on the 'DatePicker' form and selecting 'View Code'.
	HourPicker.Show 'The Date Picker form is shown and the user inputs a date.

	Call UpdateDPC(SheetName, SheetDay)
	Call Finish
End Sub

Public Function UpdateDPCByAuto()
	SheetName = Format(Date, "mmm d")
	SheetDay = Date + TimeSerial(5, 0, 0)

	DataProtector.SavePreBackup(True)

	Call Start
	Call UpdateDPC(SheetName, SheetDay, True)
	Call Finish

	UpdateDPCByAuto = DebugLogging.PrintMsg
End Function

Public Function UpdateWebBySql(SheetDate As String, Optional IsAuto As Boolean = False) As String
	Call Start
	Call WebUpdate.UpdateSql(SheetDate, IsAuto)
	Call Finish

	UpdateWebBySql = DebugLogging.PrintMsg
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------'
'The UpdateDPC subroutine is run when the Update DPC button is pressed on sheet 'Raw1'
'UpdateDPC runs the modules that load a new sheet and populate it with the requested data.
Sub UpdateDPC(SheetName As String, SheetNo As Date, Optional IsAuto As Boolean = False)
	Call CASpecific.InitializeGauges
	'-------------------------------------------------------------------------------------------------------------------------------------------------'
	'Set everything back to the way it was before starting the macro and quietly exit if the datepicker/hourpicker was closed with the close button
	If SheetName = "cancel" Then
		Call DebugLogging.PrintMsg("Cancelled.")
		Exit Sub
	End If

	Dim Answer As Integer

	Call DebugLogging.PrintMsg("Checking if Sheet Exists")

	Dim DPCsheet As Excel.Worksheet 'The 'DPCsheet' variable is used to check if the requested sheet already exists.
	'This 'For' loop checks that the a sheet for the requested date does not already exist.
	For Each DPCsheet In ThisWorkbook.Worksheets
		If DPCsheet.name = SheetName Then
			If SheetDay <> 0 And SheetNo < 42873 Then 'This If statement only matters if this was called by UpdateDPCByDate
				If Not IsAuto Then _
					MsgBox "The sheet layout changed on May 18th.  Sheets loaded before this day cannot be updated.  This update attempt will now exit."
				Call DebugLogging.PrintMsg("The sheet layout changed on May 18th.  Sheets loaded before this day cannot be updated.  This update attempt will now exit.")
				Exit Sub
			End If

			Answer = vbYes 'If the macro is being run automatically, the answer is assumed to be "Yes"
			'If a sheet with the requested date already exists, the subroutine exits so that previous data is not overwritten.
			If Not IsAuto Then _
				Answer = MsgBox("A sheet for '" & SheetName & "' already exists.  Do you want to fill the empty cells?", vbYesNo, "Sheet Already Exists")
			If Not Answer = vbYes Then
				Call DebugLogging.PrintMsg("Duplicate found.  Exiting Macro as requested by user...")
				Exit Sub
			End If

			
			Call DataProtector.EditCells(SheetName, IsAuto)
			Exit For
		End If
	Next


	'----------------------------------------------------------------------------------------------------------------------------------------------'
	'The 'AddSheet' module creates a new sheet in the workbook, names it after the requested date, and pastes the template from 'Raw2'.
	'AddSheet only runs if a sheet for the requested day does not exist.
	If Answer = 0 Then _
		Call AddSheet.CreateSheet(SheetName, SheetNo)
	'The KiWISLoader module loads the KiWIS tables to the sheet 'Raw1'.
	Call KiWISLoader.KiWIS_Import(SheetName, SheetNo, IsAuto)

	On Error Resume Next
	NextWeather = Application.WorksheetFunction.Match("Weather Forecasts:", ThisWorkbook.Sheets(SheetName).Range("A:A"), 0) + 1
	If Err.Number <> 0 Then _
		NextWeather = WeeklyStart + WeeklyCount + DataToWeatherGap
	On Error GoTo 0

	'Do not get the current forecast if today is not the same day as the one for the sheet
	If SheetName = Format(Now, "mmm d") Then _
		Call CASpecific.LoadWeather(SheetName, IsAuto)

	'The DataProtector module locks cells for editing and saves a backup of the daily planning cycle to the local desktop and the Water Management Files folder.
	Call DataProtector.LockCells(SheetName, SheetNo, IsAuto)

	Call DebugLogging.PrintMsg("Macro finished.  Setting new worksheet as active worksheet.")
	'----------------------------------------------------------------------------------------------------------------------------------------------'
	'This 'With' statement ensures that the workbook is open on the new sheet after the module has run.
	With ThisWorkbook.Worksheets(SheetName)
		'The Activate function is not recommended.  This line may cause execution errors.
		.Activate
		.Range("E16").Select
	End With

	If Not IsAuto Then _
		MsgBox "The data for " & SheetName & " has loaded."
	Call DebugLogging.PrintMsg("The data for " & SheetName & " has loaded.")
End Sub


Public Sub ChangeStatus(Optional Msg = False)
	Application.StatusBar = Msg
End Sub