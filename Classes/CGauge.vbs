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
' * @param ID   - String representing the station_name value found in Raw1 that is associated with this 
' *            Gauge.  When set to "N/A", this Gauge will not return values from Raw1.  Defaults to "N/A" if
' *             not specified.
' * @param Name - String representing the name of the gauge that is visible from the excel tables.  This value can be ignored if you are not using the WebUpdate Macros.  Defaults to the empty string.
' * 
' * @returns - This function does not return a value
' * 
' * 
' * Example usage:
' * 				Gauge.CGauge "Gauge - Widow Lake"
' * The above example sets the CGauge object Gauge's id to "Gauge - Widow Lake"
' * 
' * Example usage:
' * 				Gauge.CGauge
' * The above example sets the CGauge object Gauge's id to "N/A"
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
		If pID = "N/A" Then _
			Exit Sub
		For Each Sensor In pSensors
			If IsEmpty(.Cells(Row, Sensor.Column)) Then _
				.Cells(Row, Sensor.Column).Value = Sensor.Value(pID, IsAuto)
		Next
	End With
End Sub