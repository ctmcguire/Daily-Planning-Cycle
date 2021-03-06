Option Explicit

Private Function SendXML(xmlhttp As Object) As Boolean
	On Error GoTo OnError
	SendXML = False
	With xmlhttp
		.send
		SendXML = .waitForResponse(60000) 'This line either sets SendXML to True, sets it to False, or gets skipped, which leaves SendXML as False
		If .status <> 200 Then _
			SendXML = False
	End With
	OnError:
End Function

Private Sub TWNWeatherScraper(SheetName As String, BaseURL As String, DayOffset As Integer, Optional IsAuto As Boolean)
	'-----------------------------------------------------------------------------------------------------------------------------'
	'Please send any questions or feedback to cmcguire@mvc.on.ca
	'-----------------------------------------------------------------------------------------------------------------------------'
	''''''''''Textbox Debugger'''''''''''''
	'''''''''''''''''''''''''''''''''''''''
	'Insert these lines anywhere in the code to view the imported HTML data:
	'Debug_Text.TextBox1 = HTML_Data
	'Debug_Text.Show
	'-----------------------------------------------------------------------------------------------------------------------------'
	'The WNWeatherScraper module loads the 7 day forecast data from theweathernetwork.com.
	'-----------------------------------------------------------------------------------------------------------------------------'
	'the xmlhttp object interacts with the web server to retrieve the data
	Dim xmlhttp As Object
	'The HTML_Data variable stores all the response HTML text from the website as a string.
	Dim HTML_Data As String
	'The DataString variable is the parsed HTML code that is inserted into the Daily Planning Cycle
	Dim DataString As String
	'The TimeStamp variable are used to extract and convert the web time stamps.
	Dim WebTimeStamp As Double
	Dim DPCTimeStamp As Double
	'The Day variable is used to navigate the rows.
	Dim Day As Integer
	'Loop iterator variables
	Dim i As Integer
	Dim j As Integer

	'The Data and Column arrays are used to store values that are needed during the nested loops
	Dim Data(6) As String 'Data stores the name of the data being retrieved
	Dim Column(6) As String 'Column stores the column of the cell the retrieved data will be stored in (although for current forecast values it stores the actual cell)

	'The With statement is used to ensure the macro does not modify other workbooks that may be open.
	With ThisWorkbook
		'Initialize Data and Column arrays for the current forecast values
		Data(0) = "temperature_c"
		Data(1) = "feelsLike_c"
		Data(2) = "windDirection" '*
		Data(3) = "windGustSpeed_kmh"
		Data(4) = "name_en"

		Column(0) = "B" & DayOffset + 1
		Column(1) = "B" & DayOffset + 2
		Column(2) = "E" & DayOffset + 1
		Column(3) = "E" & DayOffset + 2
		Column(4) = "B" & DayOffset

		If .Sheets(SheetName).Range("B" & DayOffset).Value <> "" And .Sheets(SheetName).Range("B" & DayOffset).Value <> "No Response from The Weather Network" Then _
			Exit Sub

		'-----------------------------------------------------------------------------------------------------------------------------'
		''''''''''Loads the web data into VBA'''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''

		Call DebugLogging.PrintMsg("TWN - Getting weather data from server...")

		'Creates the xmlhttp object that interacts with the website. .ServerXMLHTTP60 is used so the page data is not cached.
		Set xmlhttp = New MSXML2.ServerXMLHTTP60
		'Indicates that page that will receive the request and the type of request being submitted.  Your location's link can be found at: http://legacyweb.theweathernetwork.com/rss/
		xmlhttp.Open "GET", BaseURL, True
		'Indicate that the body of the request contains form data
		xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		'Send the data as name/value pairs
		If Not SendXML(xmlhttp) Then
			Set xmlhttp = Nothing
			.Sheets(SheetName).Range("B" & DayOffset).Value = "No Response from The Weather Network"
			Exit Sub
		End If
		'Assigns the the website's HTML to the HTML_Data variable.
		HTML_Data = xmlhttp.responseText

		Call DebugLogging.PrintMsg("TWN - Finished getting weather data.  Extracting current conditions...")

		'-----------------------------------------------------------------------------------------------------------------------------'

		''''''''''Extracts the Current Conditions'''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'Isolates the observation time
		'The InStr function searches the code for the string that precedes the current conditions observation time: 'timestampApp_local'.
		'The InStr function then returns the number of characters from the start of the HTML code to the start of this string.
		'The Mid function then deletes every character before this number
		HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "timestampApp_local") + 21, Len(HTML_Data))
		WebTimeStamp = Mid(HTML_Data, 1, InStr(HTML_Data, "tzbias") - 3)
		'The website outputs the timestamp in UNIX time.  The 86,400,000 = 1000 milliseconds/second * 60 Seconds/minute * 60 Minutes/hour * 24 hours/day to convert the variable to a decimal number of days since Jan. 1, 1970.
		'The 25,569 adds the differece between Jan. 1, 1900 when Excel time starts and Jan. 1, 1970 when UNIX time begins.
		DPCTimeStamp = (WebTimeStamp / (86400000) + 25569)
		'The SheetName variable is recieved from the datepicker in the 'Update' form
		If IsEmpty(.Sheets(SheetName).Range("C" & DayOffset)) Then _
			.Sheets(SheetName).Range("C" & DayOffset).Value = DPCTimeStamp

		For j = 0 to 4 'We aren't using the whole array, so this isn't UBound(Data)
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, Chr(34) & Data(j) & Chr(34) & ":") + Len(Chr(34) & Data(j) & Chr(34) & ":"), Len(HTML_Data))
			DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1)
			DataString = Replace(DataString, Chr(34), "") 'Remove any quotation marks from the DataString

			If j = 2 Then
				DataString = Replace(DataString, "O", "W") 'Translates the wind direction to english

				'Isolates the Wind Speed
				HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "windSpeed_kmh") + 15, Len(HTML_Data))
				DataString = DataString + " " + Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1) + " km/h"
			ElseIf j = 3 Then
				DataString = DataString + " km/h"
			End If

			If IsEmpty(.Sheets(SheetName).Range(Column(j))) Then _
				.Sheets(SheetName).Range(Column(j)).Value = DataString
		next j

		Call DebugLogging.PrintMsg("TWN - Current conditions extracted.  Extracting short-term forecast...")

		'-----------------------------------------------------------------------------------------------------------------------------'

		'Initialize Data and Column arrays for the short-term forecast values (Some values will be unchanged)
		Data(2) = "pop_percent"
		Data(3) = "windDirection" '*
		Data(4) = "windGustSpeed_kmh"
		Data(5) = "rain"
		Data(6) = "snow"

		Column(0) = "C"
		Column(1) = "D"
		Column(2) = "E"
		Column(3) = "F"
		Column(4) = "G"
		Column(5) = "H"
		Column(6) = "I"

		''''''''''Extracts the Short Term Forecast'''''''''''''
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'Isolates the Short Term Forecast time
		HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "timestamp_local") + 18, Len(HTML_Data))
		WebTimeStamp = Mid(HTML_Data, 1, InStr(HTML_Data, "tzbias") - 3)
		DPCTimeStamp = (WebTimeStamp / (86400000) + 25569)
		If IsEmpty(.Sheets(SheetName).Range("B" & DayOffset + 3)) Then _
			.Sheets(SheetName).Range("B" & DayOffset + 3).Value = DPCTimeStamp

		For i = 1 to 5
			Day = DayOffset + 3 + i

			'Isolates the Forecast date
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "timestampApp_local") + 21, Len(HTML_Data))
			WebTimeStamp = Mid(HTML_Data, 1, InStr(HTML_Data, "icon") - 3)
			DPCTimeStamp = (WebTimeStamp / (86400000) + 25569)
			If IsEmpty(.Sheets(SheetName).Range("A" & Day)) Then _
				.Sheets(SheetName).Range("A" & Day).Value = DPCTimeStamp

			For j = 0 to UBound(Data)
				HTML_Data = Mid(HTML_Data, InStr(HTML_Data, Chr(34) & Data(j) & Chr(34) & ":") + Len(Chr(34) & Data(j) & Chr(34) & ":"), Len(HTML_Data))
				DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1)
				DataString = Replace(DataString, Chr(34), "") 'Remove any quotation marks from the DataString

				If j = 2 Then
					DataString = DataString + "%"
				ElseIf j = 3 Then
					DataString = Replace(DataString, "O", "W") 'Translates the wind direction to english

					'Isolates the Wind Speed
					HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "windSpeed_kmh") + 15, Len(HTML_Data))
					DataString = DataString + " " + Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1) + " km/h"
				ElseIf j = 4 Then
					DataString = DataString + " km/h"
				End If

				If IsEmpty(.Sheets(SheetName).Range(Column(j) & Day)) Then _
					.Sheets(SheetName).Range(Column(j) & Day).Value = DataString
			next j
		next i

		Call DebugLogging.PrintMsg("TWN - Short-term forecast extracted.  Extracting long-term forecast...")

		'Initialize Data and Column arrays for the long-term forecast values (Some values will be unchanged)
		Data(0) = "temperatureMin_c"
		Data(1) = "temperatureMax_c"
		Data(2) = "feelsLike_c"
		Data(4) = "popDay_percent"

		Column(0) = "D"
		Column(1) = "C"

		'-----------------------------------------------------------------------------------------------------------------------------'

		''''''''''Extracts the Long Term Forecast'''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'Isolates the Long Term Forecast time issued
		HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "timestamp_local") + 18, Len(HTML_Data))
		WebTimeStamp = Mid(HTML_Data, 1, InStr(HTML_Data, "tzbias") - 3)
		DPCTimeStamp = (WebTimeStamp / (86400000) + 25569)
		If IsEmpty(.Sheets(SheetName).Range("B" & DayOffset + 9)) Then _
			.Sheets(SheetName).Range("B" & DayOffset + 9).Value = DPCTimeStamp

		For i = 1 To 6
			Day = DayOffset + 9 + i

			'Isolates the Long Term Forecast date
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "timestampApp_local") + 21, Len(HTML_Data))
			WebTimeStamp = Mid(HTML_Data, 1, InStr(HTML_Data, "icon") - 3)
			DPCTimeStamp = (WebTimeStamp / (86400000) + 25569)
			If IsEmpty(.Sheets(SheetName).Range("A" & Day)) Then _
				.Sheets(SheetName).Range("A" & Day).Value = DPCTimeStamp

			For j = 0 To UBound(Data)
				HTML_Data = Mid(HTML_Data, InStr(HTML_Data, Chr(34) & Data(j) & Chr(34) & ":") + Len(Chr(34) & Data(j) & Chr(34) & ":"), Len(HTML_Data))
				DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1)
				DataString = Replace(DataString, Chr(34), "") 'Remove any quotation marks from the DataString

				If j = 3 Then
					DataString = Replace(DataString, "O", "W") 'Translates the wind direction to english

					'Isolates the Wind Speed
					HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "windSpeed_kmh") + 15, Len(HTML_Data))
					DataString = DataString + " " + Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1) + " km/h"
				ElseIf j = 4 Then
					DataString = DataString + "%"
				End If

				If IsEmpty(.Sheets(SheetName).Range(Column(j) & Day)) Then _
					.Sheets(SheetName).Range(Column(j) & Day).Value = DataString
			next j
		next i

		'Once the 7th day's forecast is loaded, the xmlhttp is set to 'Nothing' to prevent caching and the module closes.
		Set xmlhttp = Nothing

		Call DebugLogging.PrintMsg("TWN - Long-term forecast extracted.  Exiting macro...")
	End With
End Sub

Sub GeneralScraper(SheetName As String, LocationURL As String, Optional IsAuto As Boolean = False, Optional RowNo As Integer = 0)
	If RowNo = 0 Then _
		RowNo = NextWeather
	Call TWNWeatherScraper(SheetName, "http://legacyweb.theweathernetwork.com/dataaccess/citypage/json/" & LocationURL, RowNo, IsAuto)
	NextWeather = RowNo + TWNCount + 2
End Sub

Sub CPScraper(SheetName As String, Optional IsAuto As Boolean = False)
	Call GeneralScraper(SheetName, "caon0119", IsAuto, TWNStart)
End Sub

Sub CloyneScraper(SheetName As String, Optional IsAuto As Boolean = False)
	Call GeneralScraper(SheetName, "caon2071", IsAuto, CloyneTWNStart)
End Sub