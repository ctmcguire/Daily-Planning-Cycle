'CGaugeSensor Class
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

Private pOriginal As CGaugeSensor
Private pIsClone As Boolean

Private pInitialized As Boolean
Private pLoadedKiWIS As Boolean

Private pReturnFields As String

Private PrevVal As String


Public Sub Class_Initialize()
	pInitialized = False
	pIsClone = False
	pLoadedKiWIS = False
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
		Name = pOriginal.Name & "_timestamp"
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

Private Function FromTo(InDate As Date, Row As Integer)
	Dim FrVal As String
	Dim ToVal As String
	Dim DateVal As String
	Dim TempVal As String

	If Not pInitialized Then _
		Exit Function
	FrVal = ""
	ToVal = ""

	DateVal = Format(InDate - Switch(pIsPrev, 1, True, 0), "yyyy-mm-dd") 'Switch(cond1,value1,...) returns the first value corresponding to a condition that evaluates to true
	If Not pStartTime = "" Then
		FrVal = "&from=" & DateVal & "T" & Replace(pStartTime, "<InDate>", Hour(InDate)-pStartOffset)
		If pStartTime = "<prev>" Then
			On Error Resume Next
			Call DebugLogging.PrintMsg("Getting last staff date from " & ThisWorkbook.Sheets(Format(InDate - 1, "mmm d")).name)
			If Err.Number <> 0 Then
				FromTo = ""
				Exit Function
			End If
			On Error Goto 0
			TempVal = Format(ThisWorkbook.Sheets(Format(InDate - 1, "mmm d")).Cells(Row, "B"), "yyyy-mm-dd HH:MM:SS")
			If PrevVal = "" Then _
				PrevVal = TempVal
			if CDate(TempVal) < CDate(PrevVal) Then _
				PrevVal = TempVal
			If PrevVal = "" Then
				FromTo = ""
				Exit Function
			End If
			FrVal = "&from=" & Replace(PrevVal, " ", "T") & ".000-05:00" & "&valueorder=desc"
		End If
	End If
	If Not pEndTime = "" Then
		ToVal = "&to=" & DateVal & "T" & Replace(pEndTime, "<InDate>", Hour(InDate))
		If pEndTime = "<prev>" Then
			ToVal = "&to=" & DateVal & "T" & Hour(InDate) & ":00:00.000-05:00"
			If SheetName = Format(InDate, "mmm d") Then _
				ToVal = "&to=" & DateVal & "T23:59:59.000-05:00"
		End If
	End If

	FromTo = FrVal & ToVal
End Function

Private Function UrlKiWIS(Row As Integer)
	Dim BaseUrl As String
	Dim ReturnFields As String

	If Not pInitialized Then _
		Exit Function

	BaseUrl = "http://waterdata.quinteconservation.ca/KiWIS/KiWIS?service=kisters&type=queryServices&request=getTimeseriesValues&datasource=0&format=html&metadata=true&md_returnfields=station_name,parametertype_name&dateformat=yyyy-MM-dd%27T%27HH:mm:ss&timeseriesgroup_id="
	ReturnFields = "&returnfields=" & pReturnFields

	UrlKiWIS = BaseURL & pTsId & FromTo(SheetDay, Row)
End Function

Private Function LoadKiWIS(Row As Integer, Optional IsAuto As Boolean = False)
	Dim Url As String
	LoadKiWIS = True

	With ThisWorkbook.Sheets("Raw1").QueryTables("ExternalData_" & pRangeIndex)
		Url = "URL;" & UrlKiWIS(Row)
		If .Connection = Url Then _
			Exit Function
		.Connection = "URL;" & UrlKiWIS(Row)
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

Public Function Value(ID As String, Row As Integer, Optional IsAuto As Boolean = False)
	Dim Range As String

	If Not pInitialized Then _
		Exit Function
	'If Not pLoadedKiWIS Then
	If pOriginal.Value(ID, Row, IsAuto) = Empty Then _
		Exit Function
	'	pLoadedKiWIS = True'Make sure the original table is loaded
	'End If

	Range = GetRange()
	Value = GetData(ID, Range)
End Function

Private Function RangeHelper(Col As Integer, Range As String, Colon As Integer, Optional i As Integer = 0)
	RangeHelper = Right(Range, Len(Range) - (Colon+i))
	If Col = 0 Then _
		RangeHelper = Mid(Range, 1 + i, Colon - (1 + i))
End Function
Private Function GetRange(Optional Col As Integer = 1)
	Dim Range As String
	Dim Colon As Integer
	Dim Column As String
	Dim i As Integer
	Dim j As Integer
	
	Range = ThisWorkbook.Sheets("Raw1").Names("ExternalData_" & RangeIndex)
	Range = Replace(Range, "='Raw1'!", "")
	Range = Replace(Range, "$", "")
	Colon = InStr(Range, ":")
	
	For i = 1 To Len(RangeHelper(Col, Range, Colon))
		Column = Left(RangeHelper(Col, Range, Colon), i)
		If 0 < Val(RangeHelper(Col, Range, Colon, i)) Then _
			Exit For
		Column = ""
	Next i
	j = 0
	If Col = 0 And Column = Replace(Space(Len(Column)), " ", "Z") Then _
		j = 1
	If Col = 1 And Column = Replace(Space(Len(Column)), " ", "A") Then _
		j = 1

	GetRange = Column & Right(Range, Len(Range) - (i - j))
	If Col = 0 Then _
		GetRange = Left(Range, Colon) & Column & Right(Range, Len(Range) - (Colon + (i - j)))
	If Column = "" Then _
		GetRange = Range
End Function

Private Function GetData(ID As String, Range As String)
	Dim Raw As String
	Raw = Application.WorksheetFunction.Index(ThisWorkbook.Sheets("Raw1").Range(GetRange(0)), (Application.WorksheetFunction.Match(ID, ThisWorkbook.Sheets("Raw1").Range(Range), 0) + 5))
	If Raw <> "" Then _
		GetData = DateValue(Replace(Raw, "T", " ")) + TimeValue(Replace(Raw, "T", " "))
End Function