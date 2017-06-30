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
Public Sub CGauge(Name As String, Optional ID As String = "N/A")
	If pInitialized Then _
		Exit Sub
	pName = Name
	pID = ID
	Set pSensors = new Collection

	pInitialized = True 'Don't allow this function to be called more than once
End Sub

Public Property Get Name() As String
	Name = pName
End Property

Public Property Get ID() As String
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
	Dim Sensor As Variant
	For Each Sensor In Sensors
		pSensors.Add Sensor, Sensor.Name
	Next
End Sub

Public Function Remove(Name As String)
	pSensors.Remove(Name)
End Function

Public Sub LoadData(SheetName As String, Row As Integer)
	ThisWorkbook.Sheets(SheetName).Cells(Row, "G").Formula = GetFormula(SheetName, Row)
	If pID = "N/A" Then _
		Exit Sub
	For Each Sensor In pSensors
		If IsEmpty(ThisWorkbook.Sheets(SheetName).Cells(Row, Sensor.Column)) Then
			ThisWorkbook.Sheets(SheetName).Cells(Row, Sensor.Column).Value = Sensor.Value(pID)
		End If
	Next
End Sub

Private Function GetFormula(SheetName As String, Row As Integer)
	Dim Year As Integer
	Dim PrevYear As Integer
	Dim PrevSheet As String
	Dim Gauge As String
	Dim PrevRow As Integer

	Dim Range As String
	Dim StartPoint As Integer
	Dim EndPoint As Integer

	If Row < FlowStart Then
		GetFormula = ""
		Exit Function 'If the row is out of bounds, do nothing
	End If

	With ThisWorkbook
		Year = CInt(Format(Now, "yyyy"))
		PrevYear = CInt(Format(.Sheets(SheetName).Cells(Row, "F").Value, "yyyy"))
		PrevSheet = Format(.Sheets(SheetName).Cells(Row, "F").Value, "mmm d")
		Gauge = .Sheets(SheetName).Cells(Row, "A").Value

		If PrevYear < Year Then _
			PrevSheet = "Dec 31"
		If .Sheets(SheetName).Cells(Row, "F").Formula = "=+B" & Row & "-1" Then _
			PrevSheet = Format(Now - 1, "mmm d")
		With .Sheets(PrevSheet)
			StartPoint = Application.WorksheetFunction.Match("Staff Gauge", .Columns(1), 0)
			EndPoint = Application.WorksheetFunction.Match("Dam Operations:", .Columns(1), 0)

			If Row < DailyStart Then
				StartPoint = Application.WorksheetFunction.Match("Stream Gauge", .Columns(1), 0)
				EndPoint = Application.WorksheetFunction.Match("Lake Gauge", .Columns(1), 0)
			ElseIf Row < WeeklyStart Then
				StartPoint = Application.WorksheetFunction.Match("Lake Gauge", .Columns(1), 0)
				EndPoint = Application.WorksheetFunction.Match("Staff Gauge", .Columns(1), 0)
			End If

			Range = "A" & StartPoint & ":A" & EndPoint
			PrevRow = Application.WorksheetFunction.Match(Gauge, .Range(Range), 0) + StartPoint - 1
		End With
	End With
	GetFormula = "=INDIRECT(""'" & PrevSheet & "'!E" & PrevRow & """)"
End Function