'CGaugeSensor Class
'Implements IGaugeSensor

Private pColumn As String 'Column where this sensor's data will appear in the table
Private pRangeIndex As Integer 'Column where this sensor's data is retrieved from in raw1
Private pTsId As String

Private pStartTime As String
Private pEndTime As String
Private pStartOffset As Integer

Private pIsPrev As Boolean

Private pOriginal As CGaugeSensor
Private pIsClone As Boolean

Private pInitialized As Boolean
Private pLoadedKiWIS As Boolean

Private pReturnFields As String

Private PrevVal As String

Private pSqlCol As String
Private pSqlUpdate As Boolean


Public Sub Class_Initialize()
	pInitialized = False
	pIsClone = False
	pLoadedKiWIS = False
	pSqlCol = ""
End Sub

'/**
' * The CGaugeSensor Function is used to initialize the values in a new CGaugeSensor Object in place of its constructor.
' * This is due mostly to the fact that VBA does not support constructors with parameters, resulting in the 
' * need for this function.
' * 
' * @param Name   - The name of what this Sensor measures
' * @param Column - The letter of the column for this sensor in the dpc tables
' * @param RangeIndex - The external range that holds the data for this CGaugeSensor
' * @param TsId - The timeseries group id for this CGaugeSensor
' * @param StartTime - The minimum timestamp for the KiWIS data being retrieved.  <InDate> can be substituted for an hour to get the input date's hour.  Defaults to 5 seconds before the input hour.
' * @param StartOffset - If <InDate> is used in StartTime, this is the number of hours that will be subtracted from it.  This will not do anything if StartTime does not contain <InDate>.  Defaults to 1.
' * @param EndTime - The maximum timestamp for the KiWIS data being retrieved.  Like with StartTime, <InDate> can be used to get the input date's hour.  Defaults to 5 seconds after the input hour.
' * @param IsPrev - Whether or not this Sensor measures data from the previous day.  Will get the previous day's data if set to true.  Defaults to false.
' * @param IsSum - Whether or not this Sensor should get the sum of its returned data.  Only really meant to apply to the "rainfall to 0600" column at this point in time.  Defaults to false.
' * 
' * @returns - This function does not return anything
' * 
' * 
' * Example usage:
' * 				Sensor.CGaugeSensor "Flow Rate", "E", 3, 124004
' * The above example initializes the CGaugeSensor Sensor with a Name of "Flow Rate", sets its column to "E", sets its range index to 3 (for ExternalData_3), and sets its timeseries group id to 124004.
'**/
Public Function Init(Column As String, RangeIndex As Integer, TsId As String, _
						Optional StartTime As String = "<InDate>:59:55.000-05:00", Optional StartOffset As Integer = 1, _
						Optional EndTime As String = "<InDate>:00:05.000-05:00", _
						Optional IsPrev = False, Optional ReturnFields = "Timestamp,Value") As CGaugeSensor
	Dim temp As New CGaugeSensor
	Set Init = temp.CGaugeSensor(Column, RangeIndex, TsId, StartTime, StartOffset, EndTime, IsPrev, ReturnFields)
End Function
Public Function CGaugeSensor(Column As String, RangeIndex As Integer, TsId As String, _
						Optional StartTime As String = "<InDate>:59:55.000-05:00", Optional StartOffset As Integer = 1, _
						Optional EndTime As String = "<InDate>:00:05.000-05:00", _
						Optional IsPrev = False, Optional ReturnFields = "Timestamp,Value") As CGaugeSensor
	Set CGaugeSensor = Me

	If pInitialized Then _
		Exit Function
	pColumn = Column
	pRangeIndex = RangeIndex'Need to strip off "'Raw1'!" and the $'s
	pTsId = TsId
	
	pStartTime = StartTime
	pEndTime = EndTime
	pStartOffset = StartOffset

	pIsPrev = IsPrev

	pReturnFields = ReturnFields

	pInitialized = True
End Function

Public Function InitClone(Original As CGaugeSensor, Optional Column As String = "") As CGaugeSensor
	Dim temp As New CGaugeSensor
	Set InitClone = temp.Clone(Original, Column)
End Function
Public Function Clone(Original As CGaugeSensor, Optional Column As String = "") As CGaugeSensor
	Set Clone = Me

	If pInitialized Then _
		Exit Function
	Set pOriginal = Original
	pColumn = Column
	pIsClone = True
	
	pInitialized = True
End Function

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
				Call DebugLogging.Erred
				FromTo = ""
				Exit Function
			End If
			On Error Goto 0
			TempVal = Format(Evaluate("=MIN('" & Format(InDate - 1, "mmm d") & "'!B" & weeklyStart & ":B" & (weeklyStart + weeklyCount) & ")"), "yyyy-mm-dd HH:MM:SS")
			If TempVal = "" Then
				FromTo = ""
				Exit Function
			End If
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

Private Function LoadKiWISAsync(Row As Integer, Optional IsAuto As Boolean = False)
	Dim Url As String
	Dim strt As Date
	LoadKiWISAsync = True
	strt = Now()

	With ThisWorkbook.Sheets("Raw1").QueryTables("ExternalData_" & pRangeIndex)
		Url = "URL;" & UrlKiWIS(Row)
		If pLoadedKiWIS And .Connection = Url Then _
			Exit Function
		.Connection = Url
		On Error Resume Next
		.Refresh(True)
		If Err.Number <> 0 Then
			DebugLogging.Erred
			On Error Goto 0
			LoadKiWISAsync = False
			Exit Function
		End If
		While(Now() - strt < TimeSerial(0,0,10) And .Refreshing)
			Call DebugLogging.PrintMsg("Waiting for KiWIS response...")
		Wend
		If .Refreshing Then
			.CancelRefresh
			Call DebugLogging.PrintMsg("Response timed out")
		End If
		Call DebugLogging.PrintMsg("Loading and Copying data from KiWIS")
		pLoadedKiWIS = True
	End With
End Function

Private Function LoadKiWIS(Row As Integer, Optional IsAuto As Boolean = False)
	Dim Url As String
	LoadKiWIS = True

	With ThisWorkbook.Sheets("Raw1").QueryTables("ExternalData_" & pRangeIndex)
		Url = "URL;" & UrlKiWIS(Row)
		If pLoadedKiWIS And .Connection = Url Then _
			Exit Function
		.Connection = Url
		On Error Resume Next
		.Refresh(False)
		If Err.Number <> 0 Then
			DebugLogging.Erred
			On Error Goto 0
			LoadKiWIS = False
			Exit Function
		End If
		pLoadedKiWIS = True
	End With
End Function

Public Function Value(ID As String, Row As Integer, Optional IsAuto As Boolean = False)
	Dim Range As String

	If Not pInitialized Then _
		Exit Function
	If IsClone Then
		Value = pOriginal.Value(ID, Row, IsAuto)
		Exit Function
	End If

	If Not pLoadedKiWIS Then
		If Not LoadKiWIS(Row, IsAuto) Then _
			Exit Function
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
	On Error Goto Erred
	GetData = Application.WorksheetFunction.Index(ThisWorkbook.Sheets("Raw1").Range(Range), (Application.WorksheetFunction.Match(ID, ThisWorkbook.Sheets("Raw1").Range(Range), 0) + 5))
	Exit Function
	Erred:
	Call DebugLogging.Erred
	GetData = ""
End Function

Public Function GetSql(i As Integer, InputDate As String, ByRef Headers As Collection, ByRef Values As Collection, Optional UpdateOnly As Boolean = False)
	If pSqlCol = "" Then _
		Exit Function
	If UpdateOnly And Not(pSqlUpdate) Then _
		Exit Function
	With ThisWorkbook
		Dim Val As String
		Val = "NULL"
		If Not IsEmpty(.Sheets(InputDate).Range(pColumn & i)) Then _
			Val = .Sheets(InputDate).Range(pColumn & i)
		Headers.Add pSqlCol, pSqlCol
		Values.Add Val, pSqlCol
	End With
End Function

Public Function SqlCol(col As String, Optional update As Boolean = True) As CGaugeSensor
	Set SqlCol = Me
	pSqlCol = col
	pSqlUpdate = update
End Function