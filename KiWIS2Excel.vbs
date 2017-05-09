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
' The GaugeName array stores the site names.
Dim GaugeName(25) As String

'The Stream Gauge site names are assigned based on their order in Raw2
GaugeName(0) = "Gauge - Mississippi River below Marble Lake"
GaugeName(1) = "Gauge - Buckshot Creek near Plevna"
GaugeName(2) = "Gauge - Mississippi River at Ferguson Falls"
GaugeName(3) = "Gauge - Mississippi River at Appleton"
GaugeName(4) = "Gauge - Clyde River at Gordon Rapids"
GaugeName(5) = "Gauge - Clyde River near Lanark"
GaugeName(6) = "Gauge - Indian River near Blakeney"

GaugeName(7) = "Gauge - Carp River near Kinburn"
GaugeName(8) = "Gauge - Fall River at outlet Bennett Lake"
GaugeName(9) = "Gauge - Mississippi River at outlet Dalhousie Lake"

'The Lake Gauge site names are assigned based on their order in Raw2
GaugeName(10) = "Gauge - Shabomeka Lake"
GaugeName(11) = "Gauge - Mazinaw Lake"
GaugeName(12) = "Gauge - Kashwakamak Lake Gauge"
GaugeName(13) = "Gauge - Mississippi River at outlet Farm Lake"
GaugeName(14) = "Gauge - Mississagagon Lake"
GaugeName(15) = "Gauge - Big Gull Lake"
GaugeName(16) = "Gauge - Crotch Lake GOES"
GaugeName(17) = "Gauge - Palmerston Lake"
GaugeName(18) = "Gauge - Canonto Lake"
GaugeName(19) = "Gauge - Fall River at outlet Sharbot Lake"
GaugeName(20) = "Gauge - Lanark"
GaugeName(21) = "Gauge - Mississippi Lake"
GaugeName(22) = "Gauge - Carleton Place Dam"
GaugeName(23) = "Gauge - Poole Creek at Maple Grove"
GaugeName(24) = "Gauge - Carp River at Maple Grove"
GaugeName(25) = "Gauge - Carp River at Richardson"

'The cdpa and cdpb variables are used to calculate the current day precipitation from 0 to 6 am.
Dim cdpa As Integer
Dim cdpb As Integer
'The variables 'i' and 'z' are used as counters in the loop.
Dim i As Integer
Dim z As Integer

'The With statement is used to ensure the macro does not modify other workbooks that may be open.
With ThisWorkbook

'The 'z' variable is used to navigate the rows of the loaded sheet.
z = 6
'The 'i' counter navigates the GaugeName array.
For i = 0 To 9
'This for loop moves the Water Surveys of Canada (WSC) data from Raw1 to the loaded sheet.
'The WSC sites measure the level, flow and precipitation.
    'Inserting all the battery levels of stream guages'
    If IsEmpty(.Sheets(SheetName).Cells(z, 14)) = True Then _
    .Sheets(SheetName).Cells(z, 14).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("T1:T350"), (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("T1:T350"), 0) + 5))
    'The 6:00 am level data is extracted from Column B in Raw1.
    'Note that the .Range("B1:B350") will need to be extended if more time series are added to a group.
    'The Match function finds the correct time series in the column and the Index function returns the value.
    If IsEmpty(.Sheets(SheetName).Cells(z, 4)) = True Then _
    .Sheets(SheetName).Cells(z, 4).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("B1:B500"), (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("B1:B350"), 0) + 5))
    'The 6:00 am flow data is extracted from Column H in Raw1.
    If IsEmpty(.Sheets(SheetName).Cells(z, 5)) = True Then _
    .Sheets(SheetName).Cells(z, 5).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("H1:H500"), (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("H1:H350"), 0) + 5))
    'The previous day's precipitation data is extracted from Column Q in Raw1.
    If IsEmpty(.Sheets(SheetName).Cells(z, 11)) = True Then _
    .Sheets(SheetName).Cells(z, 11).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("E1:E500"), (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("E1:E350"), 0) + 5))

        'This If statement determines if the precipitation gauge has output a complete dataset between 00-06:00 am.
        If (.Sheets("Raw1").Range("Q" & (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("Q1:Q500"), 0) + 2))) = 7 Then
            'If the dataset is complete, the 00-06:00 am precipitation is summed and extracted.
            cdpa = (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("Q1:Q500"), 0) + 5)
            cdpb = cdpa + 12
            If IsEmpty(.Sheets(SheetName).Cells(z, 12)) = True Then _
            .Sheets(SheetName).Cells(z, 12).Value = Application.WorksheetFunction.Sum(.Sheets("Raw1").Range("Q" & cdpa, "Q" & cdpb))
            z = z + 1
        Else
            z = z + 1
        End If
Next i

'After the WSC Stream Gauge data is loaded the MVCA Lake data is loaded.
z = 21
For i = 10 To UBound(GaugeName)
    'Inserting all the battery levels of the lake gauges.
    If IsEmpty(.Sheets(SheetName).Cells(z, 14)) = True Then _
    .Sheets(SheetName).Cells(z, 14).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("T1:T350"), (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("T1:T350"), 0) + 5))
    'This if statement determines if Shabomeka Lake, Palmerston or Carp River at Maple Grove data is being loaded.
    'The Shabomeka Lake, Crotch Lake, Palmerston Lake and Carp River at Maple Grove stations measure the water level and a precipitation.
    If z = 21 Or z = 27 Or z = 28 Or z = 37 Then
        If IsEmpty(.Sheets(SheetName).Cells(z, 5)) = True Then _
        .Sheets(SheetName).Cells(z, 5).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("B1:B350"), (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("B1:B350"), 0) + 5))
        If IsEmpty(.Sheets(SheetName).Cells(z, 11)) = True Then _
        .Sheets(SheetName).Cells(z, 11).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("E1:E350"), (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("E1:E350"), 0) + 5))
            If (.Sheets("Raw1").Range("Q" & (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("Q1:Q350"), 0) + 2))) = 7 Then
                cdpa = (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("Q1:Q350"), 0) + 5)
                cdpb = cdpa + 12
                If IsEmpty(.Sheets(SheetName).Cells(z, 12)) = True Then _
                .Sheets(SheetName).Cells(z, 12).Value = Application.WorksheetFunction.Sum(.Sheets("Raw1").Range("Q" & cdpa, "Q" & cdpb))
            End If
        z = z + 1
    'This ElseIf statement determines if the Mississippi or Kashwakamak Lake data is being loaded.
    'The Mississippi and Kashwakamak Lake stations measure water levels, water temperature and air temperature.
    ElseIf z = 23 Or z = 34 Then
        If IsEmpty(.Sheets(SheetName).Cells(z, 5)) = True Then _
        .Sheets(SheetName).Cells(z, 5).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("B1:B350"), (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("B1:B350"), 0) + 5))
        If IsEmpty(.Sheets(SheetName).Cells(z, 11)) = True Then _
        .Sheets(SheetName).Cells(z, 11).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("N1:N350"), (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("N1:N350"), 0) + 5))
        z = z + 1
    'This ElseIf statement determines if the Sharbot Lake data is being loaded.
    'The Sharbot Lake WSC gauge measures the water level and precipitation.
    ElseIf z = 30 Then
        If IsEmpty(.Sheets(SheetName).Cells(z, 4)) = True Then _
        .Sheets(SheetName).Cells(z, 4).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("B1:B350"), (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("B1:B350"), 0) + 5))
        If IsEmpty(.Sheets(SheetName).Cells(z, 11)) = True Then _
        .Sheets(SheetName).Cells(z, 11).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("E1:E350"), (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("E1:E350"), 0) + 5))
            If (.Sheets("Raw1").Range("Q" & (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("Q1:Q350"), 0) + 2))) = 7 Then
                cdpa = (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("Q1:Q350"), 0) + 5)
                cdpb = cdpa + 12
                If IsEmpty(.Sheets(SheetName).Cells(z, 12)) = True Then _
                .Sheets(SheetName).Cells(z, 12).Value = Application.WorksheetFunction.Sum(.Sheets("Raw1").Range("Q" & cdpa, "Q" & cdpb))
            End If
        z = z + 3
    'The Else statement loads all other sites that measure water levels.
    Else
        If IsEmpty(.Sheets(SheetName).Cells(z, 5)) = True Then _
        .Sheets(SheetName).Cells(z, 5).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("B1:B350"), (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("B1:B350"), 0) + 5))
        z = z + 1
    End If
Next i

'The loop is reset to extract the water temperature data from Mazinaw, Kashwakamak, Mississagagon, Big Gull, Palmerston, Canonto, Lanark and Farm water temperature data.
z = 22
i = 11
Do While i < 26
    If IsEmpty(.Sheets(SheetName).Cells(z, 13)) = True Then _
    .Sheets(SheetName).Cells(z, 13).Value = Application.WorksheetFunction.Index(.Sheets("Raw1").Range("K1:K350"), (Application.WorksheetFunction.Match(GaugeName(i), .Sheets("Raw1").Range("K1:K350"), 0) + 5))
    'These inline If statements are used to skip rows in the GaugeName array and the loaded sheet for sites with no temperature data.
    If z = 26 Then z = z + 1: i = i + 1
    If z = 29 Then z = z + 3: i = i + 1
    If z = 33 Then z = z + 2: i = i + 2
    If z = 36 Then z = z + 1: i = i + 1
    z = z + 1
    i = i + 1
Loop


End With
End Sub
