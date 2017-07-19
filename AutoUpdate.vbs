Option Explicit

'-------------------------------------------------------------------------------------------------------------------------------------------------'
'The DailyUpdate subroutine is run when the Update DPC button is pressed on sheet 'Raw1'
'DailyUpdate runs the modules that load a new sheet and populate it with the requested data.
Public Function DailyUpdate()
	Call DebugLogging.PrintMsg("Creating CGauge and CGaugeSensor Objects...")
	Call CASpecific.InitializeGauges
	Call DebugLogging.PrintMsg("Finished creating CGauge and CGaugeSensor Objects.")
	'-------------------------------------------------------------------------------------------------------------------------------------------------'
	'The 'DPCsheet' variable is used to check if the requested sheet already exists.
	Dim DPCsheet As Excel.Worksheet
	Dim InDay As String
	Dim InNumber As Date

	Call DebugLogging.PrintMsg("Initializing variables...")
	'The Status Bar is located on the bottom left corner of the Excel window.  It's default status is 'READY'.
	'The Status Bar Displays 'Processing Request...' until the subroutine has ended.
	Application.StatusBar = "Processing Request..."
	'Screen Updating is turned off to speed up the processing time.
	Application.ScreenUpdating = False
	'Sheet calculations are turned off to speed up the processing time.
	Application.Calculation = xlCalculationManual

	InDay = Format(Date, "mmm d")
	InNumber = Date + TimeSerial(6, 0, 0)

	Call DebugLogging.PrintMsg("Initialization finished.  Checking if sheet exists...")
	'This 'For' loop checks that the a sheet for the requested date does not already exist.
	For Each DPCsheet In ThisWorkbook.Worksheets
		If DPCsheet.name = InDay Then
			Call DebugLogging.PrintMsg("Sheet Exists.  Missing weather data will be filled.")
			DPCsheet.Unprotect

			Call DebugLogging.PrintMsg("Getting Carleton Place forecast from AccuWeather...")
			Call WeatherAccu.CPScraper(InDay)
			Call DebugLogging.PrintMsg("Finished.  Getting Ottawa forecast from Environment Canada...")
			Call WeatherEC.OttawaScraper(InDay)
			Call DebugLogging.PrintMsg("Finished.  Getting Carleton Place forecast from The Weather Network...")
			Call WeatherTWN.CPScraper(InDay)
			Call DebugLogging.PrintMsg("Finished.  Getting Cloyne forecast from AccuWeather...")
			Call WeatherAccu.CloyneScraper(InDay)
			Call DebugLogging.PrintMsg("Finished.  Getting Cloyne forecast from The Weather Network...")
			Call WeatherTWN.CloyneScraper(InDay)

			Call DebugLogging.PrintMsg("Finished filling missing weather.  Locking cells and saving...")
			Call DataProtector.LockCells(InDay, InNumber)

			Call DebugLogging.PrintMsg("Finished locking cells and saving.  Exiting the Macro...")
			'The previously adjusted modes are returned to their default state.
			Application.StatusBar = False
			Application.Calculation = xlCalculationAutomatic
			Application.ScreenUpdating = True

			DailyUpdate = DebugLogging.PrintMsg
			Exit Function
		End If
	Next
	'----------------------------------------------------------------------------------------------------------------------------------------------'
	Call DebugLogging.PrintMsg("Worksheet not found.  Adding worksheet...")
	'The 'AddSheet' module creates a new sheet in the workbook, names it after the requested date, and pastes the template from 'Raw2'.
	Call AddSheet.CreateSheet(InDay, InNumber)
	'The KiWISLoader module loads the KiWIS tables to the sheet 'Raw1'.
	Call DebugLogging.PrintMsg("Worksheet added.  Importing data from KiWIS into Worksheet...")
	Call KiWISLoader.KiWIS_Import(InDay, InNumber, True)
	Call DebugLogging.PrintMsg("KiWIS data imported into Worksheet.  Loading weather data.")

	'The Weather... modules scrape weather data from AccuWeather, Environment Canada and The Weather Network and pastes it into the new sheet.
	Call DebugLogging.PrintMsg("Getting Carleton Place forecast from AccuWeather...")
	Call WeatherAccu.CPScraper(InDay)
	Call DebugLogging.PrintMsg("Finished.  Getting Ottawa forecast from Environment Canada...")
	Call WeatherEC.OttawaScraper(InDay)
	Call DebugLogging.PrintMsg("Finished.  Getting Carleton Place forecast from The Weather Network...")
	Call WeatherTWN.CPScraper(InDay)
	Call DebugLogging.PrintMsg("Finished.  Getting Cloyne forecast from AccuWeather...")
	Call WeatherAccu.CloyneScraper(InDay)
	Call DebugLogging.PrintMsg("Finished.  Getting Cloyne forecast from The Weather Network...")
	Call WeatherTWN.CloyneScraper(InDay)

	Call DebugLogging.PrintMsg("Finished loading weather data.  Locking cells and saving...")
	'The DataProtector module locks cells for editing and saves a backup of the daily planning cycle to the local desktop and the Water Management Files folder.
	Call DataProtector.LockCells(InDay, InNumber, True)

	Call DebugLogging.PrintMsg("Finished locking cells and saving.  Setting Worksheet to ""active""...")
	'----------------------------------------------------------------------------------------------------------------------------------------------'
	'This 'With' statement ensures that the workbook is open on the new sheet after the module has run.
	With ThisWorkbook.Worksheets(InDay)
		'The Activate function is not recommended.  This line may cause execution errors.
		.Activate
		.Range("E16").Select
	End With

	Call DebugLogging.PrintMsg("Worksheet set to ""active"".  Exiting the Macro")

	'The previously adjusted modes are returned to their default state.
	Application.StatusBar = False
	Application.Calculation = xlCalculationAutomatic
	Application.ScreenUpdating = True

	DailyUpdate = DebugLogging.PrintMsg
	Exit Function

	OnError:
		Call DebugLogging.Erred
		DailyUpdate = DebugLogging.PrintMsg
End Function