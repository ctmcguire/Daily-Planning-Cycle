Option Explicit

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
' * INSTRUCTIONS FOR ADDING A NEW GAUGE:
' * 		1.  Increase or decrease flowCount, dailyCount, and/or weeklyCount by the number of gauges being 
' * 			added to their respective CGauge arrays
' * 		3.  Use the CGauge function to initialize the new Gauges in the Gauge Arrays.  (See the CGauge 
' * 			class file for more information on how to use the CGauge.CGauge function)
' * 		4.  Use the Add function to add the desired sensors to the new Gauges.  (See the CGauge class file for
' * 			more information on how to use the CGauge.Add function)
' * 		5.  Use 'i = i + 1' to increment the i value.  
' * 		6.  Add a new row for each of the new gauges into the Raw2 table
' * 
' * INSTRUCTIONS FOR ADDING A NEW SENSOR
' * 		1.  Create a new private CGaugeSensor variable at the top of the file (with the other sensor variables)
' * 		3.  Create a new public String constant that has the variable name of the new Sensor variable prefixed 
' * 			by "Name" (take a look at the other SensorName constants if you do not understand).  This constant
' * 			should store the (unique) name of what the Sensor measures, but as long as its value is unique it 
' * 			CAN be whatever String you want without affecting the Macros (storing the name is just meant to 
' * 			help with readability)
' * 		4.  Use 'Set = New CGaugeSensor' to create the new sensor object (look at the other sensors if you 
' * 			don't know how to do this)
' * 		5.  Use the CGaugeSensor function to initialize the new sensor.  The Sensor name should be passed the 
' * 			constant you defined in step 3.  (See the CGaugeSensor class for more information on how to use 
' * 			the CGaugeSensor.CGaugeSensor function)
' * 
' * INSTRUCTIONS FOR CHANGING THE COLUMN OF A SENSOR
' * 		1.  Change the parameter for the column being changed (dpc column or raw1 column) to whichever column you wish to give it
' * 
' * INSTRUCTIONS FOR CHANGING THE ROW OF A GAUGE
' * 		1.  Because of the i variable being incremented after initializing each Gauge, the row is entirely 
' * 			dependent on where it is initialized.  So, to change its row you just need to move its 2 to 3 
' * 			lines of code to whereever you wish the Gauge to appear in the table
' * 		2.  Move the respective row in Raw2 to finish changing the Gauge's row
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


Sub LoadWeather(SheetName As String)
Call WeatherAccu.GeneralScraper(SheetName, "carleton-place/k7c/daily-weather-forecast/55438")
	Call WeatherTWN.GeneralScraper(SheetName, "caon0119")
	Call WeatherEC.GeneralScraper(SheetName, "on-118_e")
	Call WeatherAccu.GeneralScraper(SheetName, "cloyne/k0h/daily-weather-forecast/2291535")
	Call WeatherTWN.GeneralScraper(SheetName, "caon2071")
End Sub