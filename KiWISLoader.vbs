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
	'The variable 'i' is used as a counter in the loop.
	Dim i As Integer

	Dim ShouldRefresh As Boolean
	ShouldRefresh = SufficientConnections()
	'A loop is used to load the KiWIS tables into Raw1.
	'The 'i' counter navigates the TimeSeriesID array.
	For i = 0 To SensorCount - 1
		If Not ShouldRefresh Then _
			Call RecreateConnection(i)
	Next i

	Call DebugLogging.PrintMsg("KiWIS data successfully imported into Raw1.  Copying data into Worksheet...")
	Call KiWIS2Excel.Raw1Import(SheetName)
	Call DebugLogging.PrintMsg("Data copied into Worksheet.")
End Sub


Private Function SufficientConnections()
	SufficientConnections = False

	Dim qt As QueryTable
	Dim nm As Name
	Dim Range As String

	If ThisWorkbook.Sheets("Raw1").QueryTables.Count = SensorCount Then
		Call DebugLogging.PrintMsg("Number of connections is correct.  Refreshing connections...")
		SufficientConnections = True
		Exit Function
	End If

	Call DebugLogging.PrintMsg("Number of connections is incorrect.  Removing existing connections...")
	'This loop removes all QueryTable connections so as to not bog down the worksheet and/or excel file.
	For Each qt In ThisWorkbook.Sheets("Raw1").QueryTables
		qt.Delete
	Next

	Call DebugLogging.PrintMsg("Clearing and removing related external ranges...")
	For Each nm In ThisWorkbook.Sheets("Raw1").Names
		'The previously loaded data in 'Raw1' is deleted to make room for the new data.
		ThisWorkbook.Sheets("Raw1").Range(Replace(Replace(nm, "='Raw1'!", ""), "$", "")).ClearContents
		nm.Delete
	Next

	Call DebugLogging.PrintMsg("Recreating correct number of connections...")
End Function

Private Sub RecreateConnection(i As Integer)
	With ThisWorkbook.Sheets("Raw1").QueryTables.Add(Connection:="URL;", Destination:=ThisWorkbook.Sheets("Raw1").Cells(2, 3 * i + 1))
		.BackgroundQuery = True
		.TablesOnlyFromHTML = True
	End With
End Sub