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

	'These constants represent the different exception cases for the FlowGaugeExceptions array
	Const MANUAL As Integer = 0 'Set this to true if the Gauge receives its data through manual entry
	Const NO_RAIN As Integer = 1 'Set this to true if the Gauge does not measure rainfall

	'These constants represent the different exception cases for the LevelGaugeExceptions array
	Const SKIP_IT As Integer = 0 'Set this to true if the Gauge is manual, or if no data is to be entered for any other reason
	Const STAGE As Integer = 1 'Set this to true if the Gauge has a value in the Stage column
	Const AIR_TEMP As Integer = 2 'Set this to true if the Gauge tracks air temperature
	Const RAIN As Integer = 3 'Set this to true if the Gauge tracks precipitation
	Const NO_TEMP As Integer = 4 'Set this to true if the Gauge does not track water temperature

	Const FlowGaugeCount As Integer = 12
	Const LevelGaugeCount As Integer = 16

	Const FlowOffset As Integer = 6
	Const LevelOffset As Integer = FlowOffset + FlowGaugeCount + 5

	'The GaugeExceptions arrays keep track of exceptions that gauges may fall into
	Dim FlowGaugeExceptions(FlowGaugeCount, 1) As Boolean
	Dim LevelGaugeExceptions(LevelGaugeCount, 4) As Boolean 'This one needs to be 2d because it has multiple exceptions

	' The GaugeName arrays store the site names.
	Dim FlowGaugeNames(FlowGaugeCount) As String
	Dim LevelGaugeNames(LevelGaugeCount) As String

	'The cdpa and cdpb variables are used to calculate the current day precipitation from 0 to 6 am.
	Dim cdpa As Integer
	Dim cdpb As Integer
	'The variables 'i', 'j' and 'z' are used as counters in the loop.
	Dim i As Integer
	Dim j As Integer
	Dim z As Integer

	For i = 0 To UBound(FlowGaugeExceptions, 1) 'Need to specify ranks when calling UBound on a multidimensional array
		For j = 0 To UBound(FlowGaugeExceptions, 2)
			FlowGaugeExceptions(i, j) = False
		Next j
	Next i

	For i = 0 To UBound(LevelGaugeExceptions, 1) 'Need to specify ranks when calling UBound on a multidimensional array
		For j = 0 To UBound(LevelGaugeExceptions, 2)
			LevelGaugeExceptions(i, j) = False
		Next j
	Next i

	'The Stream Gauge site names are assigned based on their order in Raw2
	FlowGaugeNames(0) = "Gauge - Mississippi River below Marble Lake"
	FlowGaugeNames(1) = "Gauge - Buckshot Creek near Plevna"
	FlowGaugeNames(2) = "Gauge - Mississippi River at Ferguson Falls"
	FlowGaugeNames(3) = "Gauge - Mississippi River at Appleton"
	FlowGaugeNames(4) = "Gauge - Clyde River at Gordon Rapids"
	FlowGaugeNames(5) = "Gauge - Clyde River near Lanark"
	FlowGaugeNames(6) = "Gauge - Indian River near Blakeney"

	FlowGaugeNames(7) = "Gauge - Carp River near Kinburn"
	FlowGaugeNames(8) = "Gauge - Fall River at outlet Bennett Lake"
	FlowGaugeNames(9) = "Gauge - Mississippi River at outlet Dalhousie Lake"

	FlowGaugeNames(10) = "Gauge - Mississippi High Falls"
	FlowGaugeExceptions(10, MANUAL) = True

	FlowGaugeNames(11) = "Gauge - Poole Creek at Maple Grove"
	FlowGaugeExceptions(11, NO_RAIN) = True
	FlowGaugeNames(12) = "Gauge - Carp River at Richardson"
	FlowGaugeExceptions(12, NO_RAIN) = True


	'The Lake Gauge site names are assigned based on their order in Raw2
	LevelGaugeNames(0) = "Gauge - Shabomeka Lake"
	LevelGaugeExceptions(0, RAIN) = True
	LevelGaugeExceptions(0, NO_TEMP) = True

	LevelGaugeNames(1) = "Gauge - Mazinaw Lake"

	LevelGaugeNames(2) = "Gauge - Kashwakamak Lake Gauge"
	LevelGaugeExceptions(2, AIR_TEMP) = True

	LevelGaugeNames(3) = "Gauge - Mississippi River at outlet Farm Lake"
	LevelGaugeNames(4) = "Gauge - Mississagagon Lake"
	LevelGaugeNames(5) = "Gauge - Big Gull Lake"

	LevelGaugeNames(6) = "Gauge - Crotch Lake GOES"
	LevelGaugeExceptions(6, RAIN) = True
	LevelGaugeExceptions(6, NO_TEMP) = True

	LevelGaugeNames(7) = "Gauge - Mississippi High Falls"
	LevelGaugeExceptions(7, SKIP_IT) = True
	LevelGaugeNames(8) = "Gauge - Mississippi River at outlet Dalhousie Lake"
	LevelGaugeExceptions(8, SKIP_IT) = True

	LevelGaugeNames(9) = "Gauge - Palmerston Lake"
	LevelGaugeExceptions(9, RAIN) = True

	LevelGaugeNames(10) = "Gauge - Canonto Lake"
	LevelGaugeNames(11) = "Gauge - Lanark"

	LevelGaugeNames(12) = "Gauge - Fall River at outlet Sharbot Lake"
	LevelGaugeExceptions(12, STAGE) = True
	LevelGaugeExceptions(12, RAIN) = True
	LevelGaugeExceptions(12, NO_TEMP) = True

	LevelGaugeNames(13) = "Gauge - Fall River at outlet Bennett Lake"
	LevelGaugeExceptions(13, SKIP_IT) = True

	LevelGaugeNames(14) = "Gauge - Mississippi Lake"
	LevelGaugeExceptions(14, AIR_TEMP) = True
	LevelGaugeExceptions(14, NO_TEMP) = True

	LevelGaugeNames(15) = "Gauge - Carleton Place Dam"
	LevelGaugeExceptions(15, NO_TEMP) = True

	LevelGaugeNames(16) = "Gauge - Carp River at Maple Grove"
	LevelGaugeExceptions(16, RAIN) = True
	LevelGaugeExceptions(16, NO_TEMP) = True

	'The With statement is used to ensure the macro does not modify other workbooks that may be open.
	With ThisWorkbook

		'The 'z' variable is used to navigate the rows of the loaded sheet.
		'The 'i' counter navigates the GaugeName array.
		For i = 0 To UBound(FlowGaugeNames)
			'This for loop moves the Water Surveys of Canada (WSC) data from Raw1 to the loaded sheet.
			'The WSC sites measure the level, flow and precipitation.
			z = FlowOffset + i

			If Not FlowGaugeExceptions(i, MANUAL) Then
				'Inserting all the battery levels of stream gauges'
				If IsEmpty(.Sheets(SheetName).Cells(z, 14)) = True Then _
					.Sheets(SheetName).Cells(z, 14).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("T1:T350"), (Application.WorksheetFunction.Match(FlowGaugeNames(i), .Sheets("Raw1").Range("T1:T350"), 0) + 5))
				'The 6:00 am level data is extracted from Column B in Raw1.
				'Note that the .Range("B1:B350") will need to be extended if more time series are added to a group.
				'The Match function finds the correct time series in the column and the Index function returns the value.
				If IsEmpty(.Sheets(SheetName).Cells(z, 4)) = True Then _
					.Sheets(SheetName).Cells(z, 4).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("B1:B500"), (Application.WorksheetFunction.Match(FlowGaugeNames(i), .Sheets("Raw1").Range("B1:B350"), 0) + 5))
				'The 6:00 am flow data is extracted from Column H in Raw1.
				If IsEmpty(.Sheets(SheetName).Cells(z, 5)) = True Then _
					.Sheets(SheetName).Cells(z, 5).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("H1:H500"), (Application.WorksheetFunction.Match(FlowGaugeNames(i), .Sheets("Raw1").Range("H1:H350"), 0) + 5))
				'The previous day's precipitation data is extracted from Column E in Raw1.
				If IsEmpty(.Sheets(SheetName).Cells(z, 11)) = True And Not FlowGaugeExceptions(i, NO_RAIN) Then _
					.Sheets(SheetName).Cells(z, 11).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("E1:E500"), (Application.WorksheetFunction.Match(FlowGaugeNames(i), .Sheets("Raw1").Range("E1:E350"), 0) + 5))
				'This If statement determines if the precipitation gauge has output a complete dataset between 00-06:00 am.
				If Not FlowGaugeExceptions(i, NO_RAIN) Then
					If (.Sheets("Raw1").Range("Q" & (Application.WorksheetFunction.Match(FlowGaugeNames(i), .Sheets("Raw1").Range("Q1:Q500"), 0) + 2))) = 7 Then
						'If the dataset is complete, the 00-06:00 am precipitation is summed and extracted.
						cdpa = (Application.WorksheetFunction.Match(FlowGaugeNames(i), .Sheets("Raw1").Range("Q1:Q500"), 0) + 5)
						cdpb = cdpa + 12
						If IsEmpty(.Sheets(SheetName).Cells(z, 12)) = True Then _
							.Sheets(SheetName).Cells(z, 12).Value = Application.WorksheetFunction.Sum(.Sheets("Raw1").Range("Q" & cdpa, "Q" & cdpb))
					End If
				End If
			End If
		Next i

		'After the WSC Stream Gauge data is loaded the MVCA Lake data is loaded.
		For i = 0 To UBound(LevelGaugeNames)
			z = LevelOffset + i 'z should ideally be removed and replaced with an offset value that is added to i in the future

			If Not LevelGaugeExceptions(i, SKIP_IT) Then
				'Inserting all the battery levels of the lake gauges.
				If IsEmpty(.Sheets(SheetName).Cells(z, 14)) = True Then _
					.Sheets(SheetName).Cells(z, 14).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("T1:T350"), (Application.WorksheetFunction.Match(LevelGaugeNames(i), .Sheets("Raw1").Range("T1:T350"), 0) + 5))
				'Inserting the HG from Raw1
				If IsEmpty(.Sheets(SheetName).Cells(z, 4)) = True And LevelGaugeExceptions(i, STAGE) Then _
					.Sheets(SheetName).Cells(z, 4).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("B1:B350"), (Application.WorksheetFunction.Match(LevelGaugeNames(i), .Sheets("Raw1").Range("B1:B350"), 0) + 5))
				If IsEmpty(.Sheets(SheetName).Cells(z, 5)) = True And Not LevelGaugeExceptions(i, STAGE) Then _
					.Sheets(SheetName).Cells(z, 5).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("B1:B350"), (Application.WorksheetFunction.Match(LevelGaugeNames(i), .Sheets("Raw1").Range("B1:B350"), 0) + 5))
				
				If IsEmpty(.Sheets(SheetName).Cells(z, 11)) = True Then
					'The air temperature data is extracted from Column N in Raw1.
					If LevelGaugeExceptions(i, AIR_TEMP) Then
						.Sheets(SheetName).Cells(z, 11).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("N1:N350"), (Application.WorksheetFunction.Match(LevelGaugeNames(i), .Sheets("Raw1").Range("N1:N350"), 0) + 5))
					'The previous day's precipitation data is extracted from Column E in Raw1.
					ElseIf LevelGaugeExceptions(i, RAIN) Then
						.Sheets(SheetName).Cells(z, 11).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("E1:E350"), (Application.WorksheetFunction.Match(LevelGaugeNames(i), .Sheets("Raw1").Range("E1:E350"), 0) + 5))
					End If
				End If
				If LevelGaugeExceptions(i, RAIN) Then
					'If the dataset is complete, the 00-06:00 am precipitation is summed and extracted.
					If (.Sheets("Raw1").Range("Q" & (Application.WorksheetFunction.Match(LevelGaugeNames(i), .Sheets("Raw1").Range("Q1:Q350"), 0) + 2))) = 7 Then
						cdpa = (Application.WorksheetFunction.Match(LevelGaugeNames(i), .Sheets("Raw1").Range("Q1:Q350"), 0) + 5)
						cdpb = cdpa + 12
						If IsEmpty(.Sheets(SheetName).Cells(z, 12)) = True Then _
							.Sheets(SheetName).Cells(z, 12).Value = Application.WorksheetFunction.Sum(.Sheets("Raw1").Range("Q" & cdpa, "Q" & cdpb))
					End If
				End If
				'The water temperature data is extracted from Column K in Raw1.
				If IsEmpty(.Sheets(SheetName).Cells(z, 13)) = True And Not LevelGaugeExceptions(i, NO_TEMP) Then _
					.Sheets(SheetName).Cells(z, 13).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("K1:K350"), (Application.WorksheetFunction.Match(LevelGaugeNames(i), .Sheets("Raw1").Range("K1:K350"), 0) + 5))
			End If
		Next i

	End With
End Sub