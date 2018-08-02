'CGauge Class
Private pName As String 'Gauge name as it appears in the SQL database
Private pID As String 'Gauge name as it appears in the KiWIS data
Private pOverwrite As Boolean
Private pSensors As Collection 'Collection of Sensors.

Private pInitialized As Boolean 'Boolean value to prevent the CGauge function from being called twice on any given CGauge function


Public Sub Class_Initialize()
	pInitialized = False
	pOverwrite = False
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
Public Function Init(Optional ID As String = "N/A", Optional Name As String = "") As CGauge
	Dim temp As New CGauge
	Set Init = temp.CGauge(ID, Name)
End Function
Public Function CGauge(Optional ID As String = "N/A", Optional Name As String = "") As CGauge
	Set CGauge = Me

	If pInitialized Then _
		Exit Function
	pName = Name
	pID = ID
	Set pSensors = new Collection

	pInitialized = True 'Don't allow this function to be called more than once
End Function

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

Public Function OverwriteBlanks(Optional Val As Boolean = True) As CGauge
	Set OverwriteBlanks = Me

	pOverwrite = Val
End Function

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
Public Function Add(ParamArray Sensors() As Variant) As CGauge
	Set Add = Me

	If Not pInitialized Then _
		Exit Function
	Dim Sensor As Variant
	For Each Sensor In Sensors
		pSensors.Add Sensor, CStr(n)
	Next
End Function

Public Property Get count() As Integer
	
End Property
Private Property Get n() As Integer
	n = pSensors.Count
End Property

Private Function OverWrite(SheetName As String, Row As Integer, Col As String) As Boolean
	OverWrite = True
	With ThisWorkbook.Sheets(SheetName)
		If IsEmpty(.Cells(Row, Col)) Then _
			Exit Function
		OverWrite = .Cells(Row, Col) < CDate(SheetName)
	End With
End Function

Public Function LoadData(SheetName As String, Row As Integer, Optional IsAuto As Boolean = False) As CGauge
	Set LoadData = Me

	If Not pInitialized Then _
		Exit Function
	If n < 1 Then _
		Exit Function
	Dim updateRow As Boolean
	Dim i As Integer
	Dim keys As Collection
	Dim temp As Collection
	Set keys = new Collection
	Set temp = new Collection
	With ThisWorkbook.Sheets(SheetName)
		If pID = "N/A" Then _
			Exit Function
		updateRow = OverWrite(SheetName, Row, "B")
		i = 0
		For Each Sensor In pSensors
			keys.Add CStr(i), CStr(i) 'Need the keys later
			temp.Add .Cells(Row, Sensor.Column).Value, CStr(i) 'Need to track old values in case we need to revert back
			i = i + 1
			If ((updateRow And pOverwrite) or IsEmpty(.Cells(Row, Sensor.Column))) And Not .Cells(Row, Sensor.Column).HasFormula() Then _
				.Cells(Row, Sensor.Column).Value = Sensor.Value(pID, Row, IsAuto)
		Next
		If (.Cells(Row, "B") < DateValue(SheetName) + 1 And .Cells(Row, "B") <> "") Or Not pOverwrite Then _
			Exit Function 'If the value is not after the sheet's date, then stop here
		For Each k in keys
			.Cells(Row, pSensors.item(k).Column).Value = temp.item(k)
		Next
	End With
End Function

Public Function RunSql(i As Integer, InputDate As String, LevelsConn As ADODB.Connection) As CGauge
	Set RunSql = Me
	Dim StrSql As String
	StrSql = InsertStr(i, InputDate) & UpdateStr(i, InputDate)
	LevelsConn.Execute strSQL
End Function

Private Function InsertStr(i As Integer, InputDate As String) As String
	InsertStr = ""

	Dim havg As String
	With ThisWorkbook
		havg = "NULL"
		If Not IsEmpty(.Sheets(InputDate).Range("I" & i)) Then _
			havg = .Sheets(InputDate).Range("I" & i)
	End With

	Dim Headers As New Collection
	With Headers
		.Add "id", "id"
		.Add "gauge", "gauge"
		.Add "historicalaverage", "historicalaverage"
	End With

	Dim Values As New Collection
	With Values
		.Add "NULL", "id"
		.Add pName, "gauge"
		.Add havg, "historicalaverage"
	End With

	Dim Sensor
	For Each Sensor in pSensors
		Sensor.getSQL i, InputDate, Headers, Values
	Next

	If Headers.Count < 1 Then _
		Exit Function

	Dim HList() As String: ReDim HList(0 To Headers.Count - 1)
	Dim VList() As String: ReDim VList(0 To Headers.Count - 1)
	Dim j As Integer
	j = 0

	Dim Col
	For Each Col In Headers
		HList(j) = Col
		VList(j) = Values(Col)
		If Values(Col) <> "NULL" Then _
			VList(j) = "'" & esc(Values(Col)) & "'"
		j = j + 1
	Next

	InsertStr = "INSERT INTO mvconc55_mvclevels.data (" & Join(HList, ", ") & ") " & _
			"VALUES (" & Join(VList, ", ") & ")"
End Function
Private Function UpdateStr(i As Integer, InputDate As String) As String
	UpdateStr = ""

	Dim havg As String
	With ThisWorkbook
		havg = "NULL"
		If Not IsEmpty(.Sheets(InputDate).Range("I" & i)) Then _
			havg = .Sheets(InputDate).Range("I" & i)
	End With

	Dim Headers As New Collection
	With Headers
		.Add "historicalaverage", "historicalaverage"
	End With

	Dim Values As New Collection
	With Values
		.Add havg, "historicalaverage"
	End With

	Dim Sensor
	For Each Sensor in pSensors
		Sensor.getSQL i, InputDate, Headers, Values, True
	Next

	If Headers.Count < 1 Then _
		Exit Function
	Dim updates() As String: ReDim updates(0 To Headers.Count - 1)
	Dim j As Integer
	j = 0

	Dim Col
	For Each Col in Headers
		Dim Val As String
		Val = "NULL"
		If Values(Col) <> "NULL" Then _
			Val = "'" & esc(Values(Col)) & "'"
		updates(j) = Col & "=" & Val
		j = j + 1
	Next

	UpdateStr = " ON DUPLICATE KEY UPDATE " & Join(updates, ", ")
End Function

Private Function esc(txt As String) As String
	esc = Trim(Replace(txt, "'", "\'"))
End Function