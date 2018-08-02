Option Explicit

'/* 
' * This macro file is meant to contain the code that will be different between different CAs.  The goal is to simplify updating the DPC spreadsheets by only updating the files that are identical across all CAs.
' * other files can of course be modified, but doing so will risk the modified files needing to be re-modified whenever these files get updated.
' */

Public Const SensorCount As Integer = 9

public const flowCount as integer = 12
public const dailyCount as integer = 17
public const weeklyCount as integer = 26

'Public FlowGauges(flowCount) As CGauge
Public DailyGauges(dailyCount) As CGauge
Public WeeklyGauges(weeklyCount) As CGauge

Public Tables As CTableList

Public Const FlowOffset As Integer = 6 'Set this to the first row that gets KiWIS data
Public Const FlowToDailyGap As Integer = 5 'Set this to one more than the number of rows between the last flow gauge and first daily lake gauge
Public Const DailyToWeeklyGap As Integer = 4 'Set this to one more than the number of rows between the last daily lake gauge and first weekly lake gauge

Public Const flowStart As Integer = FlowOffset
Public Const dailyStart As Integer = flowStart + flowCount + FlowToDailyGap
Public Const weeklyStart As Integer = dailyStart + dailyCount + DailyToWeeklyGap
Public Const DataToWeatherGap As Integer = 12

Public Const LockRange As String = "A1:P" & (weeklyStart + weeklyCount + DataToWeatherGap -1)
Public Const BackupFolder As String = "\\APP-SERVER\Data_drive\common_folder\Water Management Files"

Public Const Recipients As String = "cmcguire@mvc.on.ca; gmountenay@mvc.on.ca; jnorth@mvc.on.ca; abroadbent@mvc.on.ca; jprice@mvc.on.ca; plehman@mvc.on.ca"

Private Function InitializeSensors() As Collection
	Dim fVals As New CGaugeSensor 'CGaugeSensor factory
	Dim fTime As New CTimeGaugeSensor 'CTimeGaugeSensor factory
	Dim fRmrk As New CTagGaugeSensor 'CTagGaugeSensor factory
	Dim temp As New Collection

	temp.Add fVals.Init("E", 3, 124004).SqlCol("datainfo"), "Flow"
	temp.Add fTime.InitClone(temp("Flow"), "B"), "FlowTimestamp"
	temp.Add fVals.Init("E", 1, 91667).SqlCol("datainfo"), "Level"
	temp.Add fTime.InitClone(temp("Level"), "B"), "LevelTimestamp"
	temp.Add fVals.InitClone(temp("Level"), "D"), "Stage"
	temp.Add fVals.Init("K", 2, 123967, "00:00:00.000-05:00", , "23:59:59.000-05:00", True).SqlCol("precipitation"), "Rain24H"
	temp.Add fVals.Init("L", 6, 127937), "Rain"
	temp.Add fVals.Init("K", 5, 124035), "ATemp"
	temp.Add fVals.Init("M", 4, 124025), "WTemp"
	temp.Add fVals.Init("N", 7, 291931), "Batt"
	temp.Add fVals.Init("E", 8, 574563, "<prev>", 0, "<prev>").SqlCol("datainfo"), "StaffLevel"
	temp.Add fRmrk.Init("K", 9, 574655, "<prev>", 0, "<prev>"), "StaffTag"
	temp.Add fRmrk.InitClone(temp("StaffTag"), "L"), "StaffComment"
	temp.Add fTime.InitClone(temp("StaffLevel"), "B"), "StaffTimestamp"

	Set InitializeSensors = temp
End Function

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
	Dim Sensors As Collection
	Set Sensors = InitializeSensors()

	Set Tables = New CTableList

	Dim fGauge As New CGauge 'CGauge Factory
	Dim fTable As New CTable 'CTable Factory
	With fGauge
		Tables.Add fTable.Init(6) _
			.Add(.Init("Gauge - Mississippi River below Marble Lake", "Myers Cave flow") _
				.Add(Sensors("FlowTimestamp"), Sensors("Stage"), Sensors("Flow"), Sensors("Rain24H"), Sensors("Rain"), Sensors("Batt")))
		Tables.Table().Add .Init("Gauge - Buckshot Creek near Plevna", "Buckshot Creek flow") _
			.Add(Sensors("FlowTimestamp"), Sensors("Stage"), Sensors("Flow"), Sensors("Rain24H"), Sensors("Rain"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Mississippi River at Ferguson Falls", "Ferguson Falls flow") _
			.Add(Sensors("FlowTimestamp"), Sensors("Stage"), Sensors("Flow"), Sensors("Rain24H"), Sensors("Rain"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Mississippi River at Appleton", "Appleton flow") _
			.Add(Sensors("FlowTimestamp"), Sensors("Stage"), Sensors("Flow"), Sensors("Rain24H"), Sensors("Rain"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Clyde River at Gordon Rapids", "Gordon Rapids flow") _
			.Add(Sensors("FlowTimestamp"), Sensors("Stage"), Sensors("Flow"), Sensors("Rain24H"), Sensors("Rain"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Clyde River near Lanark", "Lanark stream flow") _
			.Add(Sensors("FlowTimestamp"), Sensors("Stage"), Sensors("Flow"), Sensors("Rain24H"), Sensors("Rain"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Indian River near Blakeney", "Mill of Kintail flow") _
			.Add(Sensors("FlowTimestamp"), Sensors("Stage"), Sensors("Flow"), Sensors("Rain24H"), Sensors("Rain"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Carp River near Kinburn", "Kinburn flow") _
			.Add(Sensors("FlowTimestamp"), Sensors("Stage"), Sensors("Flow"), Sensors("Rain24H"), Sensors("Rain"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Fall River at outlet Bennett Lake", "Bennett Lake outflow") _
			.Add(Sensors("FlowTimestamp"), Sensors("Stage"), Sensors("Flow"), Sensors("Rain24H"), Sensors("Rain"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Mississippi River at outlet Dalhousie Lake", "Dalhousie Lk outflow") _
			.Add(Sensors("FlowTimestamp"), Sensors("Stage"), Sensors("Flow"), Sensors("Rain24H"), Sensors("Rain"), Sensors("Batt"))

		Tables.Table().Add .Init("Gauge - Mississippi High Falls", "High Falls Flow") _
			.Add(Sensors("FlowTimestamp"), Sensors("Flow"))

		Tables.Table().Add .Init("Gauge - Poole Creek at Maple Grove", "Poole Creek at Maple Grove") _
			.Add(Sensors("FlowTimestamp"), Sensors("Stage"), Sensors("Flow"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Carp River at Richardson", "Carp River at Richardson") _
			.Add(Sensors("FlowTimestamp"), Sensors("Stage"), Sensors("Flow"), Sensors("Batt"))
	End With

	With fGauge
		Tables.Add fTable.Init(Tables.Table().row + Tables.Table().count + 4) _
			.Add(.Init("Gauge - Shabomeka Lake", "Shabomeka Lake") _
				.Add(Sensors("LevelTimestamp"), Sensors("Level"), Sensors("Rain24H"), Sensors("Rain"), Sensors("Batt")))
		Tables.Table().Add .Init("Gauge - Mazinaw Lake", "Mazinaw Lake") _
			.Add(Sensors("LevelTimestamp"), Sensors("Level"), Sensors("WTemp"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Kashwakamak Lake Gauge", "Kashwakamak Lake") _
			.Add(Sensors("LevelTimestamp"), Sensors("Level"), Sensors("ATemp"), Sensors("WTemp"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Mississippi River at outlet Farm Lake", "Farm Lake") _
			.Add(Sensors("LevelTimestamp"), Sensors("Level"), Sensors("WTemp"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Mississagagon Lake", "Mississagagon Lake") _
			.Add(Sensors("LevelTimestamp"), Sensors("Level"), Sensors("WTemp"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Big Gull Lake", "Big Gull Lake") _
			.Add(Sensors("LevelTimestamp"), Sensors("Level"), Sensors("WTemp"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Crotch Lake GOES", "Crotch Lake") _
			.Add(Sensors("LevelTimestamp"), Sensors("Level"), Sensors("Rain24H"), Sensors("Rain"), Sensors("Batt"))

		Tables.Table().Add .Init("Gauge - Mississippi High Falls", "High Falls") _
			.Add(Sensors("LevelTimestamp"), Sensors("Level"))
		Tables.Table().Add .Init("Gauge - Mississippi River at outlet Dalhousie Lake", "Dalhousie Lake") _
			.Add(Sensors("LevelTimestamp"), Sensors("Stage"), Sensors("Level"), Sensors("Rain24H"), Sensors("Rain"), Sensors("Batt"))

		Tables.Table().Add .Init("Gauge - Palmerston Lake", "Palmerston Lake") _
			.Add(Sensors("LevelTimestamp"), Sensors("Level"), Sensors("Rain24H"), Sensors("Rain"), Sensors("WTemp"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Canonto Lake", "Canonto Lake") _
			.Add(Sensors("LevelTimestamp"), Sensors("Level"), Sensors("WTemp"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Lanark", "Lanark") _
			.Add(Sensors("LevelTimestamp"), Sensors("Level"), Sensors("WTemp"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Fall River at outlet Sharbot Lake", "Sharbot Lake") _
			.Add(Sensors("LevelTimestamp"), Sensors("Stage"), Sensors("Rain24H"), Sensors("Rain"), Sensors("Batt"))

		Tables.Table().Add .Init("Gauge - Fall River at outlet Bennett Lake", "Bennett Lake") _
			.Add(Sensors("LevelTimestamp"), Sensors("Stage"), Sensors("Level"), Sensors("Rain24H"), Sensors("Rain"), Sensors("Batt"))

		Tables.Table().Add .Init("Gauge - Mississippi Lake", "Mississippi Lake") _
			.Add(Sensors("LevelTimestamp"), Sensors("Level"), Sensors("ATemp"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Carleton Place Dam", "C.P. Dam") _
			.Add(Sensors("LevelTimestamp"), Sensors("Level"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Carp River at Maple Grove", "Carp River at Maple Grove") _
			.Add(Sensors("LevelTimestamp"), Sensors("Level"), Sensors("Rain24H"), Sensors("Rain"), Sensors("Batt"))
		Tables.Table().Add .Init("Gauge - Widow Lake", "Widow Lake") _
			.Add(Sensors("LevelTimestamp"), Sensors("Level"), Sensors("WTemp"), Sensors("Batt"))
	End With

	With fGauge
		Tables.Add fTable.Init(Tables.Table().row + Tables.Table().count + 3) _
			.Add(.Init("Gauge - Shabomeka Lake", "Shabomeka Lake (weekly)").OverwriteBlanks() _
				.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment")))
		Tables.Table().Add .Init("Gauge - Mazinaw Lake", "Mazinaw Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Little Marble Lake", "Little Marble Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Mississagagon Lake", "Mississagagon Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Kashwakamak Lake Gauge", "Kashwakamak Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Mississippi River at outlet Farm Lake", "Farm Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Mississippi River at Ardoch Bridge", "Ardoch Bridge (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Malcolm Lake", "Malcolm Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Pine Lake", "Pine Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Big Gull Lake", "Big Gull Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Buckshot Lake", "Buckshot Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Crotch Lake GOES", "Crotch Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Mississippi River at High Falls", "High Falls G.S. (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Mosque Lake", "Mosque Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Summit Lake", "Summit Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Palmerston Lake", "Palmerston Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Canonto Lake", "Canonto Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Fall River at outlet Bennett Lake", "Bennett Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Mississippi River at outlet Dalhousie Lake", "Dalhousie Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Silver Lake", "Silver Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Fall River at outlet Sharbot Lake", "Sharbot Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Widow Lake", "Widow Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Lanark", "Lanark Bridge (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Lanark Dam", "Lanark Dam (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Mississippi River at Almonte Bridge", "Almonte Bridge (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Clayton Lake", "Clayton Lake (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
		Tables.Table().Add .Init("Gauge - Carleton Place Dam", "C.P. Dam (weekly)").OverwriteBlanks() _
			.Add(Sensors("StaffTimestamp"), Sensors("StaffLevel"), Sensors("StaffTag"), Sensors("StaffComment"))
	End With
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