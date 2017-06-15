Option Explicit

Sub Raw1Import(SheetName As String)
	'-----------------------------------------------------------------------------------------------------------------------------'
	'Please send any questions or feedback to cmcguire@mvc.on.ca
	'-----------------------------------------------------------------------------------------------------------------------------'
	''''''''''Textbox Debugger'''''''''''''
	'''''''''''''''''''''''''''''''''''''''
	'Insert these lines anywhere in the code to view strings:
	'Debug_Text.TextBox1 = GaugeName(i)
	'Debug_Text.Show
	'-----------------------------------------------------------------------------------------------------------------------------'
	'The KiWIS2Excel module populates the new sheet with the values that were loaded from the KiWIS server into sheet 'Raw1'.
	'-----------------------------------------------------------------------------------------------------------------------------'

	'The variable 'i' is used as a counter in the loops.
	Dim i As Integer

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
End Sub