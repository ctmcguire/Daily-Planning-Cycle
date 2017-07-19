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
	'The variable 'i' is used as a counter in the loops.
	Dim i As Integer
	Call SufficientConnections

	Call DebugLogging.PrintMsg("Loading and Copying data from KiWIS")

	'The With statement is used to ensure the macro does not modify other workbooks that may be open.
	With ThisWorkbook
		'The 'i' counter navigates the GaugeName array.
		For i = 0 To UBound(FlowGauges)
			'This for loop moves the Water Surveys of Canada (WSC) data from Raw1 to the loaded sheet.
			'The WSC sites measure the level, flow and precipitation.
			FlowGauges(i).LoadData SheetName, i+flowStart
		Next i

		'After the WSC Stream Gauge data is loaded the MVCA Lake data is loaded.
		For i = 0 To UBound(DailyGauges)
			DailyGauges(i).LoadData SheetName, i+dailyStart
		Next i

		'After MVCA Daily Lake data is loaded, the Weekly Lake data is loaded*
		' *Currently No weekly gauges have Sensors to get data from, but this could conceivably change in the future
		For i = 0 To UBound(WeeklyGauges)
			WeeklyGauges(i).LoadData SheetName, i+weeklyStart
		Next i
	End With

	Call DebugLogging.PrintMsg("Data loaded and copied into Worksheet.")
End Sub


Private Function SufficientConnections()
	SufficientConnections = False

	If ThisWorkbook.Sheets("Raw1").QueryTables.Count = SensorCount Then
		Call DebugLogging.PrintMsg("Number of connections is correct.  Refreshing connections...")
		SufficientConnections = True
		Exit Function
	End If

	Call DebugLogging.PrintMsg("Number of connections is incorrect.  Removing existing connections...")
	'This loop removes all QueryTable connections so as to not bog down the worksheet and/or excel file.
	Dim qt As QueryTable
	For Each qt In ThisWorkbook.Sheets("Raw1").QueryTables
		qt.Delete
	Next qt

	Call DebugLogging.PrintMsg("Clearing and removing related external ranges...")
	'This loop removes all of the Defined names, in order to improve performance
	Dim nm As Name
	For Each nm In ThisWorkbook.Sheets("Raw1").Names
		'The previously loaded data in 'Raw1' is deleted to make room for the new data.
		ThisWorkbook.Sheets("Raw1").Range(Replace(Replace(nm, "='Raw1'!", ""), "$", "")).ClearContents
		nm.Delete
	Next nm

	Call DebugLogging.PrintMsg("Recreating correct number of connections...")
	'This loop re-adds the connections.  They do not start with any urls; these will be added by the CGaugeSensors before loading them
	Dim i As Integer
	For i = 0 To SensorCount - 1
		With ThisWorkbook.Sheets("Raw1").QueryTables.Add(Connection:="URL;", Destination:=ThisWorkbook.Sheets("Raw1").Cells(2, 3 * i + 1))
			.BackgroundQuery = True
			.TablesOnlyFromHTML = True
		End With
	Next i
End Function