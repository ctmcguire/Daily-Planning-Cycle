Option Explicit

'The date picker assigns a value to the public variable 'InputHour'.
'The variable is public so that it can be used by all the called functions.
Public InputHour As Date
Public InputTime As String
'-------------------------------------------------------------------------------------------------------------------------------------------------'
'The UpdateDPC subroutine is run when the Update DPC button is pressed on sheet 'Raw1'
'UpdateDPC runs the modules that load a new sheet and populate it with the requested data.
Sub UpdateHourDPC()
	'-------------------------------------------------------------------------------------------------------------------------------------------------'
	'The 'DPCsheet' variable is used to check if the requested sheet already exists.
	Dim DPCsheet As Excel.Worksheet

	'The Status Bar is located on the bottom left corner of the Excel window.  It's default status is 'READY'.
	'The Status Bar Displays 'Processing Request...' until the UpdateDPC subroutine has ended.
	Application.StatusBar = "Processing Request..."
	'Screen Updating is turned off to speed up the processing time.
	Application.ScreenUpdating = False
	'Sheet calculations are turned off to speed up the processing time.
	Application.Calculation = xlCalculationManual

	InputTime = "cancel" 'Unless the user closes the HourPicker window with the close button, this value will get changed

	'The Date Picker form is shown and the user inputs a date.
	'Instructions to install the Date Picker control can be found by right clicking on the 'DatePicker' form and selecting 'View Code'.
	HourPicker.Show

	'Set everything back to the way it was before starting the macro and quietly exit if the datepicker was closed with the close button
	If InputTime = "cancel" Then
		Application.StatusBar = False
		Application.Calculation = xlCalculationAutomatic
		Application.ScreenUpdating = True
		Exit Sub
	End If

	'This 'For' loop checks that the a sheet for the requested date does not already exist.
	For Each DPCsheet In ThisWorkbook.Worksheets
		If DPCsheet.name = InputTime Then
			'If a sheet with the requested date already exists, the subroutine exits so that previous data is not overwritten.
			MsgBox "A sheet for '" & InputHour & "' already exists."
			'The previously adjusted modes are returned to their default state.
			Application.StatusBar = False
			Application.Calculation = xlCalculationAutomatic
			Application.ScreenUpdating = True
			Exit Sub
		End If
	Next
	'----------------------------------------------------------------------------------------------------------------------------------------------'
	'The 'AddSheet' module creates a new sheet in the workbook, names it after the requested date, and pastes the template from 'Raw2'.
	Call AddSheet.CreateSheet(InputTime, InputHour)
	'The KiWISLoader module loads the KiWIS tables to the sheet 'Raw1'.
	Call KiWISLoader.KiWIS_Import(InputHour)
	'The KiWIS2Excel module loads the data from Raw1 to the new sheet that was created.
	Call KiWIS2Excel.Raw1Import(InputTime)

	'The Weather... modules scrape weather data from AccuWeather, Environment Canada and The Weather Network and pastes it into the new sheet.
	Call WeatherAccu.AccuWeatherScraper(InputTime)
	Call WeatherEC.ECWeatherScraper(InputTime)
	Call WeatherTWN.TWNWeatherScraper(InputTime)

	'The DataProtector module locks cells for editing and saves a backup of the daily planning cycle to the local desktop and the Water Management Files folder.
	Call DataProtector.LockCells(InputTime, InputHour)

	'----------------------------------------------------------------------------------------------------------------------------------------------'
	'This 'With' statement ensures that the workbook is open on the new sheet after the module has run.
	With ThisWorkbook.Worksheets(InputTime)
		'The Activate function is not recommended.  This line may cause execution errors.
		.Activate
		.Range("E16").Select
	End With

	MsgBox "The requested sheet for " & InputTime & " has loaded."

	'The previously adjusted modes are returned to their default state.
	Application.StatusBar = False
	Application.Calculation = xlCalculationAutomatic
	Application.ScreenUpdating = True

End Sub
