Option Explicit

'-------------------------------------------------------------------------------------------------------------------------------------------------'
'The DailyUpdate subroutine is run when the Update DPC button is pressed on sheet 'Raw1'
'DailyUpdate runs the modules that load a new sheet and populate it with the requested data.

Public Sub DailyUpdate()
	Call Gauges.InitializeGauges
	'-------------------------------------------------------------------------------------------------------------------------------------------------'
	'The 'DPCsheet' variable is used to check if the requested sheet already exists.
	Dim DPCsheet As Excel.Worksheet
	Dim InDay As String
	Dim InNumber As Date

	'The Status Bar is located on the bottom left corner of the Excel window.  It's default status is 'READY'.
	'The Status Bar Displays 'Processing Request...' until the subroutine has ended.
	Application.StatusBar = "Processing Request..."
	'Screen Updating is turned off to speed up the processing time.
	Application.ScreenUpdating = False
	'Sheet calculations are turned off to speed up the processing time.
	Application.Calculation = xlCalculationManual

	InDay = Format(Date, "mmm d")
	InNumber = Date + TimeSerial(6, 0, 0)

	'This 'For' loop checks that the a sheet for the requested date does not already exist.
	For Each DPCsheet In ThisWorkbook.Worksheets
		If DPCsheet.name = InDay Then
			DPCsheet.Unprotect

			Call WeatherAccu.CPScraper(InDay)
			Call WeatherEC.OttawaScraper(InDay)
			Call WeatherTWN.CPScraper(InDay)
			Call WeatherAccu.CloyneScraper(InDay)
			Call WeatherTWN.CloyneScraper(InDay)

			Call DataProtector.LockCells(InDay, InNumber)

			'The previously adjusted modes are returned to their default state.
			Application.StatusBar = False
			Application.Calculation = xlCalculationAutomatic
			Application.ScreenUpdating = True
			Exit Sub
		End If
	Next
	'----------------------------------------------------------------------------------------------------------------------------------------------'
	'The 'AddSheet' module creates a new sheet in the workbook, names it after the requested date, and pastes the template from 'Raw2'.
	Call AddSheet.CreateSheet(InDay, InNumber)
	'The KiWISLoader module loads the KiWIS tables to the sheet 'Raw1'.
	Call KiWISLoader.KiWIS_Import(InNumber)
	'The KiWIS2Excel module loads the data from Raw1 to the new sheet that was created.
	Call KiWIS2Excel.Raw1Import(InDay)

	'The Weather... modules scrape weather data from AccuWeather, Environment Canada and The Weather Network and pastes it into the new sheet.
	Call WeatherAccu.CPScraper(InDay)
	Call WeatherEC.OttawaScraper(InDay)
	Call WeatherTWN.CPScraper(InDay)
	Call WeatherAccu.CloyneScraper(InDay)
	Call WeatherTWN.CloyneScraper(InDay)

	'The DataProtector module locks cells for editing and saves a backup of the daily planning cycle to the local desktop and the Water Management Files folder.
	Call DataProtector.LockCells(InDay, InNumber)

	'----------------------------------------------------------------------------------------------------------------------------------------------'
	'This 'With' statement ensures that the workbook is open on the new sheet after the module has run.
	With ThisWorkbook.Worksheets(InDay)
		'The Activate function is not recommended.  This line may cause execution errors.
		.Activate
		.Range("E16").Select
	End With

'	Call EmailWorksheet.DailyEmail()

	'The previously adjusted modes are returned to their default state.
	Application.StatusBar = False
	Application.Calculation = xlCalculationAutomatic
	Application.ScreenUpdating = True
End Sub