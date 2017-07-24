'CGauge Class
Private pName As String 'Gauge name as it appears in the SQL database
Private pID As String 'Gauge name as it appears in the KiWIS data
Private pSensors As Collection 'Collection of Sensors.

Private pInitialized As Boolean 'Boolean value to prevent the CGauge function from being called twice on any given CGauge function


Public Sub Class_Initialize()
	pInitialized = False
End Sub

'/**
' * The CGauge Function is used to initialize the values in a new CGauge Object in place of its constructor.
' * This is due mostly to the fact that VBA does not support constructors with parameters, resulting in the 
' * need for this function.
' * 
' * @param Name - String representing the name of the gauge that is visible from the excel tables
' * @param ID   - String representing the station_name value found in Raw1 that is associated with this 
' *            Gauge.  When set to "N/A", this Gauge will not return values from Raw1.  Defaults to "N/A" if
' *             not specified.
' * 
' * @returns - This function does not return a value
' * 
' * 
' * Example usage:
' * 				'These first 2 lines are shown for context
' * 				'Dim Gauge As CGauge
' * 				'Set Gauge = New CGauge
' * 				Gauge.CGauge "Keith", "Gauge - Keith"
' * The above example sets the CGauge object Gauge's name to "Keith", and sets its id to "Gauge - Keith"
' * 
' * Example usage:
' * 				'These first 2 lines are shown for context
' * 				'Dim Gauge As CGauge
' * 				'Set Gauge = New CGauge
' * 				Gauge.CGauge "Keith (weekly)"
' * The above example sets the CGauge object Gauge's name to "Keith", and sets its id to "N/A"
'**/
Public Sub CGauge(Optional ID As String = "N/A", Optional Name As String = "")
	If pInitialized Then _
		Exit Sub
	pName = Name
	pID = ID
	Set pSensors = new Collection

	pInitialized = True 'Don't allow this function to be called more than once
End Sub

Public Property Get Name() As String
	Name = ""
	If Not pInitialized Then _
		Exit Function
	Name = pName
End Property

Public Property Get ID() As String
	ID = ""
	If Not pInitialized Then _
		Exit Function
	ID = pID
End Property

'/**
' * The Add Function is used to add new CGaugeSensor objects to this CGauge's pSensors property.  The Sensor's
' * Name value is used as the key for the Collection.
' * 
' * @param Sensors - Any number of parameters that are passed to this function.  ParamArrays need to be of type 
' *               Variant, but the intended type of any parameters passed to this function is CGaugeSensor
' * 
' * @returns - This function does not return a value
' * 
' * 
' * Example usage:
' * 				Gauge.Add Sensor
' * The above example adds the CGaugeSensor object Sensor to the CGauge object Gauge
' * 
' * Example usage:
' * 				Gauge.Add Sensor0, Sensor1,...,SensorX
' * The above example adds the CGaugeSensor objects Sensor0 through SensorX to the CGauge object Gauge (Where X is
' * some non-negative number)
'**/
Public Sub Add(ParamArray Sensors() As Variant)
	If Not pInitialized Then _
		Exit Sub
	Dim Sensor As Variant
	For Each Sensor In Sensors
		pSensors.Add Sensor, Sensor.Name
	Next
End Sub

Public Function Remove(Name As String)
	If Not pInitialized Then _
		Exit Function
	pSensors.Remove(Name)
End Function

Public Sub LoadData(SheetName As String, Row As Integer, Optional IsAuto As Boolean = False)
	If Not pInitialized Then _
		Exit Sub
	With ThisWorkbook.Sheets(SheetName)
		'.Cells(Row, "G").Formula = GetFormula(SheetName, Row) 'I'm commenting this out for now to avoid confusion caused by having the previous reading formulas being set both here and in Raw2
		If pID = "N/A" Then _
			Exit Sub
		For Each Sensor In pSensors
			If IsEmpty(.Cells(Row, Sensor.Column)) Then _
				.Cells(Row, Sensor.Column).Value = Sensor.Value(pID, IsAuto)
		Next
	End With
End Sub

'This isn't really necessary, since it could easily just get put into the Formula bar, but since it is a long formula this will probably be easier to read and understand
Private Function GetFormula(SheetName As String, Row As Integer)
	Dim PrevSheet As String 'Formula for getting the previous sheet name (Daily and Stream gauges use a static value)
	Dim StartPoint As String 'Formula for getting the start of the range of cells to look in
	Dim EndPoint As String 'Formula for getting the end of the range of cells to look in
	Dim PrevRow As String 'Formula for getting the row of the previous reading in the previous sheet

	PrevSheet = "IF(YEAR(F" & Row & ")<YEAR($B$6),""'Dec 31'!"",""'""&TEXT(F" & Row & ",""mmm d"")&""'!"")" 'Decideds whether to use the date in column F or Dec 31st
	If ThisWorkbook.Sheets(SheetName).Cells(Row, "F").Formula = "=+B" & Row & "-1" Then _
		PrevSheet = """'" & Format(CDate(SheetName) - 1, "mmm d") & "'!""" 'Daily and stream gauges are hardcoded.  This is mostly because the F column for high falls is different from the others

	StartPoint = "MATCH(""Staff Gauge"",INDIRECT(" & PrevSheet & " & ""A:A"")" 'Assume it is a weekly gauge first
	EndPoint = "MATCH(""Dam Operations:"",INDIRECT(" & PrevSheet & " & ""A:A"")"

	'If it isn't a weekly gauge, it is either a daily gauge or a stream gauge
	If Row < WeeklyStart Then
		EndPoint = StartPoint 'Weekly gauges start where daily gauges end
		StartPoint = "MATCH(""Lake Gauge"",INDIRECT(" & PrevSheet & " & ""A:A"")" 'Assume it is daily gauge next, since that let's us use the above line
	End If

	'If it isn't a weekly gauge OR a daily gauge, it must be a stream gauge
	If Row < DailyStart Then
		EndPoint = StartPoint 'Daily gauges start where stream gauges end
		StartPoint = "MATCH(""Stream Gauge"",INDIRECT(" & PrevSheet & " & ""A:A"")"
	End If

	PrevRow = "MATCH(A6,INDIRECT(" & PrevSheet & " & ""A"" & " & Startpoint & "+1 & "":A"" & " & Endpoint & "),0) + " & Startpoint & "" 'Get the row for the previous reading

	GetFormula = "=INDIRECT(" & PrevSheet & "&""E"" & " & PrevRow 'Get the previous reading
End Function