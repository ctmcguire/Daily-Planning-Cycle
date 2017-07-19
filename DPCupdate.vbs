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
	Application.StatusBar = "Processing Request..." 'The Status Bar Displays 'Processing Request...' until the UpdateDPC subroutine has ended.
	Application.ScreenUpdating = False 'Screen Updating is turned off to speed up the processing time.
	Application.Calculation = xlCalculationManual 'Sheet calculations are turned off to speed up the processing time.
End Sub

Private Sub Finish()
	'The previously adjusted modes are returned to their default state.
	Application.StatusBar = False
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
End Sub

Sub UpdateDPCByHour()
	Call Start
	SheetName = "cancel" 'If no time is chosen, this variable is assigned to exit the macro.
	'-------------------------------------------------------------------------------------------------------------------------------------------------'
	'Instructions to install the Date Picker control can be found by right clicking on the 'DatePicker' form and selecting 'View Code'.
	HourPicker.Show 'The Date Picker form is shown and the user inputs a date.

	Call UpdateDPC(SheetName, SheetDay)
End Sub

'-------------------------------------------------------------------------------------------------------------------------------------------------'
'The UpdateDPC subroutine is run when the Update DPC button is pressed on sheet 'Raw1'
'UpdateDPC runs the modules that load a new sheet and populate it with the requested data.
Sub UpdateDPC(SheetName As String, SheetNo As Date)
	Call CASpecific.InitializeGauges
	'-------------------------------------------------------------------------------------------------------------------------------------------------'
	'Set everything back to the way it was before starting the macro and quietly exit if the datepicker/hourpicker was closed with the close button
	If SheetName = "cancel" Then
		Call Finish
		Exit Sub
	End If
	Dim Answer As Integer

	Dim DPCsheet As Excel.Worksheet 'The 'DPCsheet' variable is used to check if the requested sheet already exists.
	'This 'For' loop checks that the a sheet for the requested date does not already exist.
	For Each DPCsheet In ThisWorkbook.Worksheets
		If DPCsheet.name = SheetName Then
			If (Not SheetDay = 0) And SheetNo < 42873 Then 'This If statement only matters if this was called by UpdateDPCByDate
				MsgBox "The sheet layout changed on May 18th.  Sheets loaded before this day cannot be updated.  This update attempt will now exit."
				Call Finish
				Exit Sub
			End If
			'If a sheet with the requested date already exists, the subroutine exits so that previous data is not overwritten.
			Answer = MsgBox("A sheet for '" & SheetName & "' already exists.  Do you want to fill the empty cells?", vbYesNo, "Sheet Already Exists")
			If Not Answer = vbYes Then
				Call Finish
				Exit Sub
			End If
			DPCsheet.Unprotect
		End If
	Next


	'----------------------------------------------------------------------------------------------------------------------------------------------'
	'The 'AddSheet' module creates a new sheet in the workbook, names it after the requested date, and pastes the template from 'Raw2'.
	'AddSheet only runs if a sheet for the requested day does not exist.
	If Answer = 0 Then Call AddSheet.CreateSheet(SheetName, SheetNo)
	'The KiWISLoader module loads the KiWIS tables to the sheet 'Raw1'.
	Call KiWISLoader.KiWIS_Import(SheetName, SheetNo)

	'The Weather... modules scrape weather data from AccuWeather, Environment Canada and The Weather Network and pastes it into the new sheet.
	NextWeather = WeeklyStart + WeeklyCount + DataToWeatherGap
	If Answer = 0 Then
		Call CASpecific.LoadWeather(SheetName)
	End If

	'The DataProtector module locks cells for editing and saves a backup of the daily planning cycle to the local desktop and the Water Management Files folder.
	Call DataProtector.LockCells(SheetName, SheetNo)

	'----------------------------------------------------------------------------------------------------------------------------------------------'
	'This 'With' statement ensures that the workbook is open on the new sheet after the module has run.
	With ThisWorkbook.Worksheets(SheetName)
		'The Activate function is not recommended.  This line may cause execution errors.
		.Activate
		.Range("E16").Select
	End With

	MsgBox "The data for " & SheetName & " has loaded."

	'The previously adjusted modes are returned to their default state.
	Application.StatusBar = False
	Application.Calculation = xlCalculationAutomatic
	Application.ScreenUpdating = True

End Sub