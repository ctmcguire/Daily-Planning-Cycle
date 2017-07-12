Option Explicit

Sub KiWIS_Import(SheetName As String, InputDate As Date, Optional IsAuto As Boolean = False)
	'-----------------------------------------------------------------------------------------------------------------------------'
	'Please send any questions or feedback to cmcguire@mvc.on.ca
	'-----------------------------------------------------------------------------------------------------------------------------'
	''''''''''Textbox Debugger'''''''''''''
	'''''''''''''''''''''''''''''''''''''''
	'Insert these lines anywhere in the code to view strings:
	'Debug_Text.TextBox1 = URL
	'Debug_Text.Show
	'-----------------------------------------------------------------------------------------------------------------------------'
	'The KiWISLoader module loads the html tables from the KiWIS server.
	'The tables are loaded into sheet 'Raw1'.
	'-----------------------------------------------------------------------------------------------------------------------------'
	'The PrevDate variable is used to load the 24 hr precipitation data from the previous day.
	Dim PrevDate As String
	'The KiWISDate variable stores the requested date in a KiWIS friendly format.
	Dim KiWISDate As String
	'The URL1 and URL2 variables are used to assemble the KiWIS URL address.
	Dim URL1 As String
	Dim URL2 As String
	'The variable 'i' is used as a counter in the loop.
	Dim i As Integer
	'The variable 'qt' is used to delete the QueryTables connections.
	Dim qt As QueryTable
	'All of the time series with identical parameters are stored in groups so that the data is loaded efficiently.
	'A time series can be added to a group in WISKi by right clicking on the time series and clicking 'Add to group'.
	'KiWIS uses time series group IDs to identify the required group.
	'The Hub's Time Series Group IDs can be found at: http://waterdata.quinteconservation.ca/KiWIS/KiWIS?service=kisters&type=queryServices&request=getgrouplist&datasource=0&format=html&metadata=true&md_returnfields=station_name,parametertype_name
	'The TimeSeriesID array stores the time series group IDs.
	Dim TimeSeriesID(6) As String
	TimeSeriesID(0) = 91667 'Daily Levels
	TimeSeriesID(1) = 123967 '24h Precip
	TimeSeriesID(2) = 124004 'Daily Flows
	TimeSeriesID(3) = 124025 'Daily Water Temperature
	TimeSeriesID(4) = 124035 'Daily Air Temperature
	TimeSeriesID(5) = 127937 'Current Day Precipitation to 0600
	TimeSeriesID(6) = 291931 'Battery levels'

	'The Format function rearrages the date so that it can be processed by the KiWIS server.
	KiWISDate = Format(InputDate, "yyyy-mm-dd")
	KiWISDate = "&from=" & KiWISDate & "T" & Hour(InputDate) - 1 & ":59:55.000-05:00&to=" & KiWISDate & "T" & Hour(InputDate) & ":00:05.000-05:00"

	'The DateAdd function is nested inside the Format function and is used to calculate the previous date.
	'The Format function rearrages the date so that it can be processed by the KiWIS server.
	PrevDate = Format(DateAdd("d", -1, InputDate), "yyyy-mm-d")
	PrevDate = "&from=" & PrevDate & "T00:00:00.000-05:00&to=" & PrevDate & "T23:59:59.000-05:00"

	'The previously loaded data in 'Raw1' is deleted to make room for the new data.
	ThisWorkbook.Sheets("Raw1").Range("A2:T500").ClearContents

	'The base URL address is assigned to URL1.
	URL1 = "http://waterdata.quinteconservation.ca/KiWIS/KiWIS?service=kisters&type=queryServices&request=getTimeseriesValues&datasource=0&format=html&metadata=true&md_returnfields=station_name,parametertype_name&dateformat=yyyy-MM-dd%27T%27HH:mm:ss&timeseriesgroup_id="

	'A loop is used to load the KiWIS tables into Raw1.
	'The 'i' counter navigates the TimeSeriesID array.
	For i = 0 To UBound(TimeSeriesID)

		'Set URL2
		URL2 = URL1 & TimeSeriesID(i) & KiWISDate
		'Previous 24 h precipitation is loaded using PrevDate
		If i = 1 Then
			URL2 = URL1 & TimeSeriesID(i) & PrevDate
		'00-06:00 precipitation is loaded using a modified date.
		ElseIf i = 5 Then
			URL2 = URL1 & TimeSeriesID(i) & Left(KiWISDate, 16) & "T" & Hour(InputDate) - 6 & ":00:00.000-05:00&to=" & Right(Left(KiWISDate, 16), 10) & "T" & Hour(InputDate) & ":00:00.000-05:00"
		End If

		If Not ThisWorkbook.Sheets("Raw1").QueryTables("ExternalData_" & i+1).Connection = "URL;" & URL2 Then _
			ThisWorkbook.Sheets("Raw1").QueryTables("ExternalData_" & i+1).Connection = "URL;" & URL2
		ThisWorkbook.Sheets("Raw1").QueryTables("ExternalData_" & i+1).Refresh(False)
'		With ThisWorkbook.Sheets("Raw1").QueryTables.Add(Connection:="URL;" & URL2, Destination:=ThisWorkbook.Sheets("Raw1").Cells(2, 3 * i + 1))
'			.BackgroundQuery = True
'			.TablesOnlyFromHTML = True
'			On Error Resume Next
'			.Refresh BackgroundQuery:=False
'			If Err.Number <> 0 Then
'				On Error Goto 0
'				If Not IsAuto Then _
'					MsgBox "KiWIS Loader has failed"
'				Goto TheEnd
'			End If
'			On Error Goto 0
'			.SaveData = True
'		End With
	Next i

	Call DebugLogging.PrintMsg("KiWIS data successfully imported into Raw1.  Copying data into Worksheet...")
	Call KiWIS2Excel.Raw1Import(SheetName)
	Call DebugLogging.PrintMsg("Data copied into Worksheet.")
	TheEnd:

'	If ThisWorkbook.Sheets("Raw1").QueryTables.Count = 7 Then _
'		Exit Function

'	Call DebugLogging.PrintMsg("Removing connections...")
'	'This loop removes all QueryTable connections so as to not bog down the worksheet and/or excel file.
'	For Each qt In ThisWorkbook.Sheets("Raw1").QueryTables
'		qt.Delete
'	Next

'	Call DebugLogging.PrintMsg("Removing related defined names...")
'	Dim nm As Name
'	For Each nm In ThisWorkbook.Sheets("Raw1").Names
'		nm.Delete
'	Next

'	Call DebugLogging.PrintMsg("Defined names removed.")
End Sub