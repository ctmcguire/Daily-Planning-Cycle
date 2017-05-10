Option Explicit

Sub KiWIS_Import(InputDate As Date)
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
Dim KiWISDate2 As String
'The URL1 and URL2 variables are used to assemble the KiWIS URL address.
Dim URL1 As String
Dim URL2 As String
'The variables 'i' and 'z' are used as counters in the loop.
Dim i As Integer
Dim z As Integer
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
TimeSeriesID(5) = 291931 'Battery levels'
TimeSeriesID(6) = 127937 'Current Day Precipitation to 0600

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
URL1 = "URL;http://waterdata.quinteconservation.ca/KiWIS/KiWIS?service=kisters&type=queryServices&request=getTimeseriesValues&datasource=0&format=html&metadata=true&md_returnfields=station_name,parametertype_name&dateformat=yyyy-MM-dd%27T%27HH:mm:ss&timeseriesgroup_id="

'A loop is used to load the 6 Time Series Groups.
z = 1
'The 'i' counter navigates the TimeSeriesID array.
'Calls battery levels before previous 6 hours of precipitation which modifies KiWISDate
For i = 0 To UBound(TimeSeriesID)
    If i = 1 Then
        'Loads precipitation from previous 24h
        URL2 = URL1 & TimeSeriesID(i) & PrevDate
        'The QueryTables function downloads and imports the data from the KiWIS server to Raw1.
        With ThisWorkbook.Sheets("Raw1").QueryTables.Add(Connection:=URL2, Destination:=ThisWorkbook.Sheets("Raw1").Range("D2"))
            .BackgroundQuery = True
            .TablesOnlyFromHTML = True
            .Refresh BackgroundQuery:=False
            .SaveData = True
        End With
    ElseIf i < 5 Then
        'Loads the Daily Levels, Daily Flows, Daily Water Temperature and Daily Air Temperature time series groups.
        URL2 = URL1 & TimeSeriesID(i) & KiWISDate
        With ThisWorkbook.Sheets("Raw1").QueryTables.Add(Connection:=URL2, Destination:=ThisWorkbook.Sheets("Raw1").Cells(2, z))
            .BackgroundQuery = True
            .TablesOnlyFromHTML = True
            .Refresh BackgroundQuery:=False
            .SaveData = True
        End With
     ElseIf i = 5 Then
        'Loads the battery levels of gauges'
        URL2 = URL1 & TimeSeriesID(i) & KiWISDate
        With ThisWorkbook.Sheets("Raw1").QueryTables.Add(Connection:=URL2, Destination:=ThisWorkbook.Sheets("Raw1").Range("S2"))
          .BackgroundQuery = True
          .TablesOnlyFromHTML = True
          .Refresh BackgroundQuery:=False
          .SaveData = True
        End With
    ElseIf i = 6 Then
        'Loads Previous 6 hours of Precipitation
        KiWISDate = Left(KiWISDate, 16)
        KiWISDate = KiWISDate & "T" & Hour(InputDate) - 6 & ":00:00.000-05:00&to=" & Right(KiWISDate, 10) & "T" & Hour(InputDate) & ":00:00.000-05:00"
        URL2 = URL1 & TimeSeriesID(i) & KiWISDate
        With ThisWorkbook.Sheets("Raw1").QueryTables.Add(Connection:=URL2, Destination:=ThisWorkbook.Sheets("Raw1").Range("P2"))
            .BackgroundQuery = True
            .TablesOnlyFromHTML = True
            .Refresh BackgroundQuery:=False
            .SaveData = True
        End With
    End If

'The 'z' counter navigates the columns in Raw1.
'Each time series group takes up 3 columns.
z = z + 3
Next i

End Sub


