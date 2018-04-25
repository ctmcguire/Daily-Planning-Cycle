'CSumGaugeSensor Class
'Implements IGaugeSensor

Private pName As String 'What value the gauge sensor measures (flow, level, precipitation, etc)
Private pColumn As String 'Column where this sensor's data will appear in the table
Private pRangeIndex As Integer 'Column where this sensor's data is retrieved from in raw1
Private pTsId As String

Private pStartTime As String
Private pEndTime As String
Private pStartOffset As Integer

Private pIsPrev As Boolean
Private pIsSum As Boolean 'Whether or not this Sensor is a summation of values

Private pOriginal As CSumGaugeSensor
Private pIsClone As Boolean

Private pInitialized As Boolean
Private pLoadedKiWIS As Boolean


Public Sub Class_Initialize()
	pInitialized = False
	pIsClone = False
	pLoadedKiWIS = False
End Sub

Public Sub CSumGaugeSensor(Name As String, Column As String, RangeIndex As Integer, TsId As String, _
						Optional StartTime As String = "<InDate>:59:55.000-05:00", Optional StartOffset As Integer = 1, _
						Optional EndTime As String = "<InDate>:00:05.000-05:00", _
						Optional IsPrev = False)
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

Public Sub Clone(Original As CSumGaugeSensor, Optional Column As String = "")
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

Public Property Get RangeIndex()
	If pIsClone Then
		RangeIndex = pOriginal.RangeIndex
		Exit Function
	End If
	RangeIndex = pRangeIndex
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

	With ThisWorkbook.Sheets("Raw1").QueryTables("ExternalData_" & pRangeIndex)
		.Connection = "URL;" & UrlKiWIS()
		On Error Resume Next
		.Refresh(False)
		If Err.Number <> 0 Then
			DebugLogging.Erred
			On Error Goto 0
			LoadKiWIS = False
			Exit Function
		End If
	End With
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
			GetData = ""
			Exit Function
		End If

		Dim Row As Integer
		Row = Application.WorksheetFunction.Match(ID, .Range(Range), 0) + 6
		GetData = Application.WorksheetFunction.Sum(.Range(Column & Row, Column & Row+12))
	End With
End Function