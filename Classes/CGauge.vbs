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
	If pID = "N/A" Then _
		Exit Sub
	For Each Sensor In pSensors
		If IsEmpty(ThisWorkbook.Sheets(SheetName).Cells(Row, Sensor.Column)) Then
			ThisWorkbook.Sheets(SheetName).Cells(Row, Sensor.Column).Value = Sensor.Value(pID)
		End If
	Next
End Sub