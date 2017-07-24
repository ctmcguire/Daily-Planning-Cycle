'CGaugeSensor Class
Private pName As String 'What value the gauge sensor measures (flow, level, precipitation, etc)
Private pColumn As String 'Column where this sensor's data will appear in the table
Private pRangeIndex As Integer 'Column where this sensor's data is retrieved from in raw1
Private pTsId As String

Private pStartTime As String
Private pEndTime As String
Private pStartOffset As Integer

Private pIsPrev As Boolean
Private pIsSum As Boolean 'Whether or not this Sensor is a summation of values

Private pOriginal As CGaugeSensor
Private pIsClone As Boolean

Private pInitialized As Boolean
Private pLoadedKiWIS As Boolean


Public Sub Class_Initialize()
	pInitialized = False
	pIsClone = False
	pLoadedKiWIS = False
End Sub

' * The CGaugeSensor Function is used to initialize the values in a new CGaugeSensor Object in place of its constructor.
' * This is due mostly to the fact that VBA does not support constructors with parameters, resulting in the 
' * need for this function.
' * 
' * @param Name   - The name of what this Sensor measures
' * @param Column - The letter of the column for this sensor in the dpc tables
' * @param RawCol - The letter of the column for this sensor in the raw1 table
' * 
' * @returns - This function does not return anything
' * 
' * 
' * Example usage:
' * 				'These first 2 lines are shown for context
' * 				'Dim Sensor As CGaugeSensor
' * 				'Set Sensor = New CGaugeSensor
' * 				Sensor.CGaugeSensor "Dave the intern", "D", "I"
' * The above example initializes the CGaugeSensor Sensor with a Name of "Dave the intern", a Column of "D", and a RawCol of "I"
'**/
Public Sub CGaugeSensor(Name As String, Column As String, RangeIndex As Integer, TsId As String, _
						Optional StartTime As String = "<InDate>:59:55.000-05:00", Optional StartOffset As Integer = 1, _
						Optional EndTime As String = "<InDate>:00:05.000-05:00", _
						Optional IsPrev = False, Optional IsSum = False)
	If pInitialized Then _
		Exit Sub
	pName = Name
	pColumn = Column
	pRangeIndex = RangeIndex'Need to strip off "'Raw1'!" and the $'s
	pTsId = TsId
	
	pStartTime = StartTime
	pEndTime = EndTime
	pStartOffset = StartOffset

	pIsPrev = IsPrev
	pIsSum = IsSum

	pInitialized = True
End Sub

Public Sub Clone(Original As CGaugeSensor, Optional Column As String = "")
	If pInitialized Then _
		Exit Sub
	Set pOriginal = Original
	pColumn = Column
	pIsClone = True
	
	pInitialized = True
End Sub

Public Property Get IsClone()
	IsClone = pIsClone
End Property

Public Property Get Column()
	If pIsClone And pColumn = "" Then
		Column = pOriginal.Column
		Exit Function
	End If
	Column = pColumn
End Property

Public Property Get Name()
	If pIsClone Then
		Name = pOriginal.Name
		Exit Function
	End If
	Name = pName
End Property

Private Function FromTo(InDate As Date)
	Dim FrVal As String
	Dim ToVal As String
	Dim DateVal As String

	If Not pInitialized Then _
		Exit Function

	DateVal = Format(InDate - Switch(pIsPrev, 1, True, 0), "yyyy-mm-dd") 'Switch(cond1,value1,...) returns the first value corresponding to a condition that evaluates to true
	FrVal = "&from=" & DateVal & "T" & Replace(pStartTime, "<InDate>", Hour(InDate)-pStartOffset)
	ToVal = "&to=" & DateVal & "T" & Replace(pEndTime, "<InDate>", Hour(InDate))

	FromTo = FrVal & ToVal
End Function

Private Function UrlKiWIS()
	Dim BaseUrl As String

	If Not pInitialized Then _
		Exit Function

	BaseUrl = "http://waterdata.quinteconservation.ca/KiWIS/KiWIS?service=kisters&type=queryServices&request=getTimeseriesValues&datasource=0&format=html&metadata=true&md_returnfields=station_name,parametertype_name&dateformat=yyyy-MM-dd%27T%27HH:mm:ss&timeseriesgroup_id="

	UrlKiWIS = BaseURL & pTsId & FromTo(SheetDay)
End Function

Private Function LoadKiWIS(Optional IsAuto As Boolean = False)
	LoadKiWIS = True

	ThisWorkbook.Sheets("Raw1").QueryTables("ExternalData_" & pRangeIndex).Connection = "URL;" & UrlKiWIS()
	On Error Resume Next
	ThisWorkbook.Sheets("Raw1").QueryTables("ExternalData_" & pRangeIndex).Refresh(False)
	If Err.Number <> 0 Then
		On Error Goto 0
		If Not IsAuto Then _
			MsgBox "KiWIS Loader has failed"
		LoadKiWIS = False
		Exit Function
	End If
End Function

Public Function Value(ID As String, Optional IsAuto As Boolean = False)
	Dim Range As String

	If Not pInitialized Then _
		Exit Function
	If IsClone Then
		Value = pOriginal.Value(ID)
		Exit Function
	End If

	If Not pLoadedKiWIS Then
		If Not LoadKiWIS(IsAuto) Then _
			Exit Function
		pLoadedKiWIS = True
	End If

	Range = GetRange()

	If pIsSum Then
		Value = Sum(ID, Range)
		Exit Function
	End If
	Value = GetData(ID, Range)
End Function

Private Function GetRange()
	Dim Range As String
	Dim Colon As Integer
	Dim Column As String
	Dim i As Integer
	
	Range = ThisWorkbook.Sheets("Raw1").Names("ExternalData_" & pRangeIndex)
	Range = Replace(Range, "='Raw1'!", "")
	Range = Replace(Range, "$", "")
	Colon = InStr(Range, ":")
	
	For i = 1 To Len(Right(Range, Len(Range) - Colon))
		Column = Left(Right(Range, Len(Range) - Colon), i)
		If 0 < Val(Right(Range, Len(Range) - (Colon + i))) Then _
			Exit For
		Column = ""
	Next i
	
	GetRange = Column & Right(Range, Len(Range) - i)
	If Column = "" Then _
		GetRange = Range
End Function

Private Function GetData(ID As String, Range As String)
	GetData = Application.WorksheetFunction.Index(ThisWorkbook.Sheets("Raw1").Range(Range), (Application.WorksheetFunction.Match(ID, ThisWorkbook.Sheets("Raw1").Range(Range), 0) + 5))
End Function

Private Function Sum(ID As String, Range As String)
	Dim Column As String
	Dim i As Integer
	For i = 1 To Len(Range)
		Column = Left(Range, i)
		If 0 < Val(Column) Then _
			Exit For
		Column = Left(Range, 1)
	Next i
	
	With ThisWorkbook.Sheets("Raw1")
		If Not (.Range(Column & (Application.WorksheetFunction.Match(ID, .Range(Range), 0) + 3))) = 7 Then
			Sum = ""
			Exit Function
		End If

		Dim Row As Integer
		Row = Application.WorksheetFunction.Match(ID, .Range(Range), 0) + 6
		Sum = Application.WorksheetFunction.Sum(.Range(Column & Row, Column & Row+12))
	End With
End Function