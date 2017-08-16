Option Explicit

Private Function SendXML(xmlhttp As Object) As Boolean
	On Error GoTo OnError
	SendXML = False
	With xmlhttp
		.send
		SendXML = .waitForResponse(60000) 'This line either sets SendXML to True, sets it to False, or gets skipped, which leaves SendXML as False
	End With
	OnError:
End Function

Private Sub ECWeatherScraper(SheetName As String, BaseURL As String, DayOffset As Integer, Optional IsAuto As Boolean = False)
	'-----------------------------------------------------------------------------------------------------------------------------'
	'Please send any questions or feedback to cmcguire@mvc.on.ca
	'-----------------------------------------------------------------------------------------------------------------------------'
	''''''''''Textbox Debugger'''''''''''''
	'''''''''''''''''''''''''''''''''''''''
	'Insert these lines anywhere in the code to view the imported HTML data:
	'Debug_Text.TextBox1 = HTML_Data
	'Debug_Text.Show
	'-----------------------------------------------------------------------------------------------------------------------------'
	'The ECWeatherScraper module loads the 7 day forecast data from Environment Canada (weather.gc.ca).
	'-----------------------------------------------------------------------------------------------------------------------------'
	'the xmlhttp object interacts with the web server to retrieve the data
	Dim xmlhttp As Object
	'The HTML_Data variable stores all the response HTML text from the website as a string.
	Dim HTML_Data As String
	'The DataString variable is the parsed HTML code that is inserted into the Daily Planning Cycle
	Dim DataString As String
	Dim Low As Integer
	'The Day variable is used to navigate the rows.
	Dim Day As Integer

	Dim i As Integer 'Loop iterator variable

	Dim Data(9) As String 'The Data array stores the names for all the data so that they can be found in the html code
	Dim Cell(9) As String 'The Cell array stores the cells which are given the data from the html code

	Dim StrEnd As Integer 'The StrEnd variable is used to change where the datastring should end, for when it is desirable to cut off some of the data

	'The With statement is used to ensure the macro does not modify other workbooks that may be open.
	With ThisWorkbook
		'-----------------------------------------------------------------------------------------------------------------------------'

		Data(0) = "Observed at:"
		Cell(0) = "B" & (DayOffset + 1)

		Data(1) = "Condition:"
		Cell(1) = "B" & (DayOffset + 2)

		Data(2) = "Temperature:"
		Cell(2) = "B" & (DayOffset + 3)

		Data(3) = "Pressure / Tendency:"
		Cell(3) = "E" & (DayOffset + 3)

		Data(4) = "Visibility:"
		Cell(4) = "E" & (DayOffset + 4)

		Data(5) = "Humidity:"
		Cell(5) = "E" & (DayOffset + 5)

		Data(6) = "Wind Chill:"
		Cell(6) = "B" & (DayOffset + 4)

		Data(7) = "Dewpoint:"
		Cell(7) = "H" & (DayOffset + 3)

		Data(8) = "Wind:"
		Cell(8) = "H" & (DayOffset + 4)

		Data(9) = "Air Quality Health Index:"
		Cell(9) = "B" & (DayOffset + 5)

		If .Sheets(SheetName).Range("B" & (DayOffset)).Value <> "" And .Sheets(SheetName).Range("B" & (DayOffset)).Value <> "No Response from Environment Canada" Then _
			Exit Sub

		Call DebugLogging.PrintMsg("EC - Getting weather data from server...")

		''''''''''Loads the web data into VBA'''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''
		'Creates the xmlhttp object that interacts with the website. .ServerXMLHTTP60 is used so the page data is not cached.
		'Loads the weather page and saves the HTML data as the variable HTML_Data
		Set xmlhttp = New MSXML2.ServerXMLHTTP60
		'Indicates that page that will receive the request and the type of request being submitted.
		'Your location's link can be found by searching for a local forecast at: http://weather.gc.ca/canada_e.html
		'After the local forecast has loaded, click on the RSS Weather link underneath the historical data and adjacent to the 'Follow:" text.
		xmlhttp.Open "GET", BaseURL, True
		'Indicate that the body of the request contains form data
		xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		'Send the data as name/value pairs
		If Not SendXML(xmlhttp) Then
			Set xmlhttp = Nothing
			.Sheets(SheetName).Range("B" & (DayOffset)).Value = "No Response from Environment Canada"
			Exit Sub
		End If
		'Assigns the the website's HTML to the HTML_Data variable.
		HTML_Data = xmlhttp.responseText

		Call DebugLogging.PrintMsg("EC - Finished getting weather data from server.  Extracting current conditions...")

		'-----------------------------------------------------------------------------------------------------------------------------'

		''''''''''Extracts the Current Conditions and Watches and Warnings'''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'The irrelevant HTML code is cut.
		'The InStr function searches the code for the string that precedes the relevant data: '<rights>Copyright 2016, Environment Canada</rights>'.
		'The InStr function then returns the number of characters from the start of the HTML code to the start of this string.
		'The Mid function then deletes every character before this number
		If InStr(HTML_Data, ", Environment Canada</rights>") > 0 Then
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, ", Environment Canada</rights>"), Len(HTML_Data))
		End If

		'Isolates the watches and warnings string.
		HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<title>") + 7, Len(HTML_Data))
		DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "</title>") - 1)
		'The SheetName variable is recieved from the datepicker in the 'Update' form.
		If IsEmpty(.Sheets(SheetName).Range("B" & (DayOffset))) Then _
			.Sheets(SheetName).Range("B" & (DayOffset)).Value = DataString

		For i = 0 to UBound(Data)
			DataString = "N/A" 'Default value in case some of the data (specifically wind chill) isn't in the html string

			'If the data is in the html string, extract it
			If InStr(HTML_Data, "<b>" & Data(i) & "</b>") > 0 Then
				HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<b>" & Data(i) & "</b>") + Len("<b>" & Data(i) & "</b>"), Len(HTML_Data))
				StrEnd = 1 'Don't include the "<" from "<br/>" in the data

				If i = 2 Or i = 7 Then
					StrEnd = 8 'Remove the "degrees Celsius" from the end of the temperature and dewpoint data
				ElseIf i = 9 Then
					StrEnd = 2 'Remove that random extra space from the end of the air quality data
				End If

				DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "<br/>") - StrEnd) 'extract the data from the html string
			End If

			If IsEmpty(.Sheets(SheetName).Range(Cell(i))) Then _
				.Sheets(SheetName).Range(Cell(i)).Value = DataString 'Set the value of the appropriate cell
		next i

		Call DebugLogging.PrintMsg("EC - Current conditions extracted.  Extracting long-term forecast...")

		'-----------------------------------------------------------------------------------------------------------------------------'

		''''''''''Extracts the Long Term Forecast'''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''

		For i = 1 to 13
			Day = DayOffset + 6 + i

			'Isolates the 7 day forecast day.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<title>") + 7, Len(HTML_Data))
			DataString = Mid(HTML_Data, 1, InStr(HTML_Data, ":"))
			If IsEmpty(.Sheets(SheetName).Range("A" & Day)) Then _
				.Sheets(SheetName).Range("A" & Day).Value = DataString
			'Isolates the 7 day forecast data.
			'Chr(34) returns a double quotation mark (") and is used to prevent runtime errors.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<summary type=" & Chr(34) & "html" & Chr(34) & ">") + 21, Len(HTML_Data))
			DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "Forecast issued") - 1)
			If IsEmpty(.Sheets(SheetName).Range("B" & Day)) Then _
				.Sheets(SheetName).Range("B" & Day).Value = DataString
		next i

		'Isolates the Long Term Forecast time issued.
		HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "Forecast") + 0, Len(HTML_Data))
		DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "</summary>") - 1)
		If IsEmpty(.Sheets(SheetName).Range("A" & DayOffset + 6)) Then _
			.Sheets(SheetName).Range("A" & DayOffset + 6).Value = DataString

		'Once the 7th day's forecast is loaded, the xmlhttp is set to 'Nothing' to prevent caching and the module closes.
		Set xmlhttp = Nothing
	End With

	Call DebugLogging.PrintMsg("EC - Long-term forecast extracted.  Exiting macro...")

End Sub

Sub GeneralScraper(SheetName As String, LocationURL As String, Optional IsAuto As Boolean = False, Optional RowNo As Integer = 0)
	If RowNo = 0 Then _
		RowNo = NextWeather
	Call ECWeatherScraper(SheetName, "http://weather.gc.ca/rss/city/" & LocationURL & ".xml", RowNo, IsAuto)
	NextWeather = RowNo + ECCount + 2
End Sub

Sub OttawaScraper(SheetName As String, Optional IsAuto As Boolean = False)
	Call GeneralScraper(SheetName, "on-118_e", IsAuto, ECStart)
End Sub