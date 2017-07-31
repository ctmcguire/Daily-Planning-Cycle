Option Explicit

'/* 
' * This macro file is meant to contain the code that will be different between different CAs.  The goal is to simplify updating the DPC spreadsheets by only updating the files that are identical across all CAs.
' * other files can of course be modified, but doing so will risk the modified files needing to be re-modified whenever these files get updated.
' */

Public Const SensorCount As Integer = 7

Private Stage As CGaugeSensor
Private Flow As CGaugeSensor
Private Level As CGaugeSensor
Private Rain24H As CGaugeSensor
Private Rain As CGaugeSensor
Private ATemp As CGaugeSensor
Private WTemp As CGaugeSensor
Private Batt As CGaugeSensor

Public Const StageName As String = "Stage"
Public Const FlowName As String = "Flow"
Public Const LevelName As String = "Level"
Public Const Rain24HName As String = "Precipitation (last 24 hours)"
Public Const RainName As String = "Precipitation (to 0600)"
Public Const ATempName As String = "Air Temperature"
Public Const WTempName As String = "Water Temperature"
Public Const BattName As String = "Battery Level"

public const flowCount as integer = 12
public const dailyCount as integer = 17
public const weeklyCount as integer = 26

Public FlowGauges(flowCount) As CGauge
Public DailyGauges(dailyCount) As CGauge
Public WeeklyGauges(weeklyCount) As CGauge

Public Const FlowOffset As Integer = 6 'Set this to the first row that gets KiWIS data
Public Const FlowToDailyGap As Integer = 5 'Set this to one more than the number of rows between the last flow gauge and first daily lake gauge
Public Const DailyToWeeklyGap As Integer = 4 'Set this to one more than the number of rows between the last daily lake gauge and first weekly lake gauge

Public Const flowStart As Integer = FlowOffset
Public Const dailyStart As Integer = flowStart + flowCount + FlowToDailyGap
Public Const weeklyStart As Integer = dailyStart + dailyCount + DailyToWeeklyGap
Public Const DataToWeatherGap As Integer = 12


'/**
' * The InitializeGauges function is used to initialize the 3 CGauge array constants, and should be called
' *  at the beginning of the first Sub or Function called in a macro (currently DPCUpdate.UpdateDPC and 
' * WebUpdate.Run_WebUpdate)
' * 
' * NOTE: All changes to the DPC spreadsheet layout will require Raw2 to be modified.  Do these modifications before following the steps below
' * 
' * INSTRUCTIONS FOR ADDING A NEW SENSOR:
' * 		1.  Create a new private CGaugeSensor variable at the top of the file (with the other sensor variables)
' * 		2.  Use 'Set = New CGaugeSensor' to create the new sensor object (look at the other sensors if you 
' * 			don't know how to do this)
' * 		3.  Use the CGaugeSensor function to initialize the new sensor.  The Sensor name should reflect the information it retrieves, but it can be anything so long as its name is unique.  
' * 			(See the CGaugeSensor class for more information on how to use the CGaugeSensor.CGaugeSensor function)
' * 
' * INSTRUCTIONS FOR CHANGING A SENSOR'S COLUMN:
' * 	Changing a sensor's column is simple
' * 		1.  Find the line where the CGaugeSensor for the sensor you would like to change the column of calls its CGaugeSensor function.
' * 		2.  Change the column letter to that of the column you would like its data to appear in (this is the second parameter of the CGaugeSensor function)
' * 
' * INSTRUCTIONS FOR SETTING UP A SENSOR IN 2 COLUMNS:
' * 	Somtimes you may want a Sensor's data to appear in one column for some gauges and a different column for other gauges.  Doing this is similar (but not identical) to adding a brand new sensor.
' * 		1.  Follow steps 1 and 2 for adding a new sensor, but stop before step 3
' * 		2.  Instead of calling the new sensor's CGaugeSensor function, call its Clone function.  Pass the Gauge you want to appear in multiple columns as the first parameter, and the new column it should appear in as
' * 			the second.  Any gauge that uses the original column should use the original sensor, while gauges using the new column should use its clone.
' * 
' * INSTRUCTIONS FOR ADDING A NEW GAUGE:
' * 		1.  Increase or decrease flowCount, dailyCount, and/or weeklyCount by the number of gauges being 
' * 			added to their respective CGauge arrays
' * 		2.  Use the CGauge function to initialize the new Gauges in the Gauge Arrays.  (See the CGauge 
' * 			class file for more information on how to use the CGauge.CGauge function)
' * 		3.  Use the Add function to add the desired sensors to the new Gauges.  (See the CGauge class file for
' * 			more information on how to use the CGauge.Add function)
' * 		4.  Use 'i = i + 1' to increment the i value.  
' * 
' * INSTRUCTIONS FOR CHANGING THE ROW OF A GAUGE
' * 	The i variable is incremented after initializing each gauge.  As such, its row is based on when it was initialized.
' * 		1.  Locate the CGauge whose row you would like to change.
' * 		2.  Locate the CGauge who appears in the row just before the new row you would like the CGauge from step 1 to appear on.
' * 		3.  Cut all 2 to 3 lines for the CGauge in step 1 and paste them after the 2 to 3 lines for the CGauge in step 2.
' * 
' * INSTRUCTIONS FOR USING FEWER THAN 3 LISTS OF GAUGES
' * 	You may not use 3 groups of gauges; while there is currently no way to add additional lists (without modifying other files), it is possible to use fewer.
' * 		1.  Decide which of the lists (FlowGauges/DailyGauges/WeeklyGauges) will be ignored.  Set its respective count constant to 0 and remove all the lines in which its Gauges call their CGauge function
' * 		2.  Set the respective gap variable (FlowToDailyGap/DailyToWeeklyGap/DataToWeatherGap) to 0.
'**/
Sub InitializeGauges()
	Set Flow = New CGaugeSensor
	Flow.CGaugeSensor FlowName, "E", 3, 124004
	
	Set Level = New CGaugeSensor
	Level.CGaugeSensor LevelName, "E", 1, 91667

	Set Stage = New CGaugeSensor
	Stage.Clone Level, "D"

	Set Rain24H = New CGaugeSensor
	Rain24H.CGaugeSensor Rain24HName, "K", 2, 123967, "00:00:00.000-05:00", , "23:59:59.000-05:00", True

	Set Rain = New CGaugeSensor
	Rain.CGaugeSensor RainName, "L", 6, 127937, "<InDate>:00:00.000-05:00", 6, "<InDate>:00:00.000-05:00", , True

	Set ATemp = New CGaugeSensor
	ATemp.CGaugeSensor ATempName, "K", 5, 124035

	Set WTemp = New CGaugeSensor
	WTemp.CGaugeSensor WTempName, "M", 4, 124025

	Set Batt = New CGaugeSensor
	Batt.CGaugeSensor BattName, "N", 7, 291931


	Dim i As Integer
	For i = 0 To flowCount
		Set FlowGauges(i) = New CGauge
	Next i
	For i = 0 To dailyCount
		Set DailyGauges(i) = New CGauge
	Next i
	For i = 0 To weeklyCount
		Set WeeklyGauges(i) = New CGauge
	Next i


	i = 0
	FlowGauges(i).CGauge "Gauge - Mississippi River below Marble Lake", "Myers Cave flow"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Gauge - Buckshot Creek near Plevna", "Buckshot Creek flow"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Gauge - Mississippi River at Ferguson Falls", "Ferguson Falls flow"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Gauge - Mississippi River at Appleton", "Appleton flow"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Gauge - Clyde River at Gordon Rapids", "Gordon Rapids flow"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Gauge - Clyde River near Lanark", "Lanark stream flow"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Gauge - Indian River near Blakeney", "Mill of Kintail flow"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Gauge - Carp River near Kinburn", "Kinburn flow"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Gauge - Fall River at outlet Bennett Lake", "Bennett Lake outflow"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Gauge - Mississippi River at outlet Dalhousie Lake", "Dalhousie Lk outflow"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Gauge - Mississippi High Falls", "High Falls Flow"
	i = i + 1

	FlowGauges(i).CGauge "Gauge - Poole Creek at Maple Grove", "Poole Creek at Maple Grove"
	FlowGauges(i).Add Stage, Flow, Batt
	i = i + 1

	FlowGauges(i).CGauge "Gauge - Carp River at Richardson", "Carp River at Richardson"
	FlowGauges(i).Add Stage, Flow, Batt
	i = i + 1


	i = 0
	DailyGauges(i).CGauge "Gauge - Shabomeka Lake", "Shabomeka Lake"
	DailyGauges(i).Add Level, Rain24H, Rain, Batt
	i = i + 1

	DailyGauges(i).CGauge "Gauge - Mazinaw Lake", "Mazinaw Lake"
	DailyGauges(i).Add Level, WTemp, Batt
	i = i + 1

	DailyGauges(i).CGauge "Gauge - Kashwakamak Lake Gauge", "Kashwakamak Lake"
	DailyGauges(i).Add Level, ATemp, WTemp, Batt
	i = i + 1

	DailyGauges(i).CGauge "Gauge - Mississippi River at outlet Farm Lake", "Farm Lake"
	DailyGauges(i).Add Level, WTemp, Batt
	i = i + 1

	DailyGauges(i).CGauge "Gauge - Mississagagon Lake", "Mississagagon Lake"
	DailyGauges(i).Add Level, WTemp, Batt
	i = i + 1

	DailyGauges(i).CGauge "Gauge - Big Gull Lake", "Big Gull Lake"
	DailyGauges(i).Add Level, WTemp, Batt
	i = i + 1

	DailyGauges(i).CGauge "Gauge - Crotch Lake GOES", "Crotch Lake"
	DailyGauges(i).Add Level, Rain24H, Rain, Batt
	i = i + 1

	DailyGauges(i).CGauge "Gauge - Mississippi High Falls", "High Falls"
	i = i + 1

	DailyGauges(i).CGauge "Gauge - Mississippi River at outlet Dalhousie Lake", "Dalhousie Lake"
	i = i + 1

	DailyGauges(i).CGauge "Gauge - Palmerston Lake", "Palmerston Lake"
	DailyGauges(i).Add Level, Rain24H, Rain, WTemp, Batt
	i = i + 1

	DailyGauges(i).CGauge "Gauge - Canonto Lake", "Canonto Lake"
	DailyGauges(i).Add Level, WTemp, Batt
	i = i + 1

	DailyGauges(i).CGauge "Gauge - Lanark", "Lanark"
	DailyGauges(i).Add Level, WTemp, Batt
	i = i + 1

	DailyGauges(i).CGauge "Gauge - Fall River at outlet Sharbot Lake", "Sharbot Lake"
	DailyGauges(i).Add Stage, Rain24H, Rain, Batt
	i = i + 1

	DailyGauges(i).CGauge "Gauge - Fall River at outlet Bennett Lake", "Bennett Lake"
	i = i + 1

	DailyGauges(i).CGauge "Gauge - Mississippi Lake", "Mississippi Lake"
	DailyGauges(i).Add Level, ATemp, Batt
	i = i + 1

	DailyGauges(i).CGauge "Gauge - Carleton Place Dam", "C.P. Dam"
	DailyGauges(i).Add Level, Batt
	i = i + 1

	DailyGauges(i).CGauge "Gauge - Carp River at Maple Grove", "Carp River at Maple Grove"
	DailyGauges(i).Add Level, Rain24H, Rain, Batt
	i = i + 1

	DailyGauges(i).CGauge "Gauge - Widow Lake", "Widow Lake"
	i = i + 1


	i = 0
	WeeklyGauges(i).CGauge Name:="Shabomeka Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Mazinaw Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Little Marble Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Mississagagon Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Kashwakamak Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Farm Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Ardoch Bridge (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Malcolm Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Pine Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Big Gull Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Buckshot Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Crotch Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="High Falls G.S. (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Mosque Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Summit Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Palmerston Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Canonto Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Bennett Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Dalhousie Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Silver Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Sharbot Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Widow Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Lanark Bridge (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Lanark Dam (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Almonte Bridge (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="Clayton Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge Name:="C.P. Dam (weekly)"
	i = i + 1
End Sub


'/**
' * INSTRUCTIONS FOR CALLING THE WEATHER FUNCTIONS
' * 	Not all CAs get the same weather information; not only does the geographic location change, but the source sights aren't necessarily the same either.
' * 		1.  Decide which source or sources you will be getting weather from (the options are AccuWeather, Environment Canada, and The Weather Network).
' * 		2.  Decide which order you want your sources to appear in.  The order you call the scraper functions in WILL be the order in which the weather data appears, so keep this in mind when calling them.
' * 		3.  For each source, you will need to unique part of the url that distinguishes it from the other locations.  For the weather network, this is everything after the last forward slash.  For environment canada,
' * 			this is everything after the last forward slash but before the '.xml'.  For AccuWeather, this is everything after "www.accuweather.com/en/ca/".  
' * 		4.  Call the GeneralScraper function from the respective source's VBA module (WeatherAccu/WeatherEC/WeatherTWN), passing SheetName as parameter 1, the unique part of the url obtained in step 3 as parameter 2, 
' * 			and the IsAuto variable as parameter 3.
' * 				>  IsAuto is just a boolean that indicates whether or not UI popups should be displayed.  The popups interfere when macros are being run in the background from task scheduler
' * 		5.  Repeat steps 1 through 4 for each scraper you would like to add, keeping in mind that the order they are called in is the order they appear in.
'**/
Sub LoadWeather(SheetName As String, Optional IsAuto As Boolean = False)
	'The Weather... modules scrape weather data from AccuWeather, Environment Canada and The Weather Network and pastes it into the new sheet.
	Call WeatherAccu.GeneralScraper(SheetName, "carleton-place/k7c/daily-weather-forecast/55438", IsAuto)
	Call WeatherTWN.GeneralScraper(SheetName, "caon0119", IsAuto)
	Call WeatherEC.GeneralScraper(SheetName, "on-118_e", IsAuto)
	Call WeatherAccu.GeneralScraper(SheetName, "cloyne/k0h/daily-weather-forecast/2291535", IsAuto)
	Call WeatherTWN.GeneralScraper(SheetName, "caon2071", IsAuto)
End Sub