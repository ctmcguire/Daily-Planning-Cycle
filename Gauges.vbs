Option Explicit

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
	Set Stage = New CGaugeSensor
	Stage.CGaugeSensor StageName, "D", 1

	Set Flow = New CGaugeSensor
	Flow.CGaugeSensor FlowName, "E", 3
	
	Set Level = New CGaugeSensor
	Level.CGaugeSensor LevelName, "E", 1

	Set Rain24H = New CGaugeSensor
	Rain24H.CGaugeSensor Rain24HName, "K", 2

	Set Rain = New CGaugeSensor
	Rain.CGaugeSensor RainName, "L", 6

	Set ATemp = New CGaugeSensor
	ATemp.CGaugeSensor ATempName, "K", 5

	Set WTemp = New CGaugeSensor
	WTemp.CGaugeSensor WTempName, "M", 4

	Set Batt = New CGaugeSensor
	Batt.CGaugeSensor BattName, "N", 7


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
	FlowGauges(i).CGauge "Myers Cave flow", "Gauge - Mississippi River below Marble Lake"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Buckshot Creek flow", "Gauge - Buckshot Creek near Plevna"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Ferguson Falls flow", "Gauge - Mississippi River at Ferguson Falls"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Appleton flow", "Gauge - Mississippi River at Appleton"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Gordon Rapids flow", "Gauge - Clyde River at Gordon Rapids"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Lanark stream flow", "Gauge - Clyde River near Lanark"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Mill of Kintail flow", "Gauge - Indian River near Blakeney"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Kinburn flow", "Gauge - Carp River near Kinburn"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Bennett Lake outflow", "Gauge - Fall River at outlet Bennett Lake"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "Dalhousie Lk outflow", "Gauge - Mississippi River at outlet Dalhousie Lake"
	FlowGauges(i).Add Stage, Flow, Rain24H, Rain, Batt
	i = i + 1

	FlowGauges(i).CGauge "High Falls Flow", "Gauge - Mississippi High Falls"
	i = i + 1

	FlowGauges(i).CGauge "Poole Creek at Maple Grove", "Gauge - Poole Creek at Maple Grove"
	FlowGauges(i).Add Stage, Flow, Batt
	i = i + 1

	FlowGauges(i).CGauge "Carp River at Richardson", "Gauge - Carp River at Richardson"
	FlowGauges(i).Add Stage, Flow, Batt
	i = i + 1


	i = 0
	DailyGauges(i).CGauge "Shabomeka Lake", "Gauge - Shabomeka Lake"
	DailyGauges(i).Add Level, Rain24H, Rain, Batt
	i = i + 1

	DailyGauges(i).CGauge "Mazinaw Lake", "Gauge - Mazinaw Lake"
	DailyGauges(i).Add Level, WTemp, Batt
	i = i + 1

	DailyGauges(i).CGauge "Kashwakamak Lake", "Gauge - Kashwakamak Lake Gauge"
	DailyGauges(i).Add Level, ATemp, WTemp, Batt
	i = i + 1

	DailyGauges(i).CGauge "Farm Lake", "Gauge - Mississippi River at outlet Farm Lake"
	DailyGauges(i).Add Level, WTemp, Batt
	i = i + 1

	DailyGauges(i).CGauge "Mississagagon Lake", "Gauge - Mississagagon Lake"
	DailyGauges(i).Add Level, WTemp, Batt
	i = i + 1

	DailyGauges(i).CGauge "Big Gull Lake", "Gauge - Big Gull Lake"
	DailyGauges(i).Add Level, WTemp, Batt
	i = i + 1

	DailyGauges(i).CGauge "Crotch Lake", "Gauge - Crotch Lake GOES"
	DailyGauges(i).Add Level, Rain24H, Rain, Batt
	i = i + 1

	DailyGauges(i).CGauge "High Falls", "Gauge - Mississippi High Falls"
	i = i + 1

	DailyGauges(i).CGauge "Dalhousie Lake", "Gauge - Mississippi River at outlet Dalhousie Lake"
	i = i + 1

	DailyGauges(i).CGauge "Palmerston Lake", "Gauge - Palmerston Lake"
	DailyGauges(i).Add Level, Rain24H, Rain, WTemp, Batt
	i = i + 1

	DailyGauges(i).CGauge "Canonto Lake", "Gauge - Canonto Lake"
	DailyGauges(i).Add Level, WTemp, Batt
	i = i + 1

	DailyGauges(i).CGauge "Lanark", "Gauge - Lanark"
	DailyGauges(i).Add Level, WTemp, Batt
	i = i + 1

	DailyGauges(i).CGauge "Sharbot Lake", "Gauge - Fall River at outlet Sharbot Lake"
	DailyGauges(i).Add Stage, Rain24H, Rain, Batt
	i = i + 1

	DailyGauges(i).CGauge "Bennett Lake", "Gauge - Fall River at outlet Bennett Lake"
	i = i + 1

	DailyGauges(i).CGauge "Mississippi Lake", "Gauge - Mississippi Lake"
	DailyGauges(i).Add Level, ATemp, Batt
	i = i + 1

	DailyGauges(i).CGauge "C.P. Dam", "Gauge - Carleton Place Dam"
	DailyGauges(i).Add Level, Batt
	i = i + 1

	DailyGauges(i).CGauge "Carp River at Maple Grove", "Gauge - Carp River at Maple Grove"
	DailyGauges(i).Add Level, Rain24H, Rain, Batt
	i = i + 1

	DailyGauges(i).CGauge "Widow Lake", "Gauge - Widow Lake"
	i = i + 1


	i = 0
	WeeklyGauges(i).CGauge "Shabomeka Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Mazinaw Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Little Marble Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Mississagagon Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Kashwakamak Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Farm Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Ardoch Bridge (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Malcolm Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Pine Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Big Gull Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Buckshot Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Crotch Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "High Falls G.S. (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Mosque Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Summit Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Palmerston Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Canonto Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Bennett Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Dalhousie Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Silver Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Sharbot Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Widow Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Lanark Bridge (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Lanark Dam (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Almonte Bridge (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "Clayton Lake (weekly)"
	i = i + 1

	WeeklyGauges(i).CGauge "C.P. Dam (weekly)"
	i = i + 1
End Sub