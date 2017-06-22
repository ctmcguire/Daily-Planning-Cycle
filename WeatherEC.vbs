Option Explicit

Private Function SendXML(xmlhttp As Object) As Integer
	On Error Resume Next
	xmlhttp.send
	If Err.Number <> 0 Then
		SendXML = 1
		Exit Function
	End If
	SendXML = 0
End Function

Private Sub ECWeatherScraper(SheetName As String, BaseURL As String, DayOffset As Integer)
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
		Cell(0) = "B" & (DayOffset - 5)

		Data(1) = "Condition:"
		Cell(1) = "B" & (DayOffset - 4)

		Data(2) = "Temperature:"
		Cell(2) = "B" & (DayOffset - 3)

		Data(3) = "Pressure / Tendency:"
		Cell(3) = "E" & (DayOffset - 3)

		Data(4) = "Visibility:"
		Cell(4) = "E" & (DayOffset - 2)

		Data(5) = "Humidity:"
		Cell(5) = "E" & (DayOffset - 1)

		Data(6) = "Wind Chill:"
		Cell(6) = "B" & (DayOffset - 2)

		Data(7) = "Dewpoint:"
		Cell(7) = "H" & (DayOffset - 3)

		Data(8) = "Wind:"
		Cell(8) = "H" & (DayOffset - 2)

		Data(9) = "Air Quality Health Index:"
		Cell(9) = "B" & (DayOffset - 1)

		If .Sheets(SheetName).Range("B" & (DayOffset - 6)).Value <> "" And .Sheets(SheetName).Range("B" & (DayOffset - 6)).Value <> "No Response from Environment Canada" Then _
			Exit Sub

		''''''''''Loads the web data into VBA'''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''
		'Creates the xmlhttp object that interacts with the website. .ServerXMLHTTP60 is used so the page data is not cached.
		'Loads the weather page and saves the HTML data as the variable HTML_Data
		Set xmlhttp = New MSXML2.ServerXMLHTTP60
		'Indicates that page that will receive the request and the type of request being submitted.
		'Your location's link can be found by searching for a local forecast at: http://weather.gc.ca/canada_e.html
		'After the local forecast has loaded, click on the RSS Weather link underneath the historical data and adjacent to the 'Follow:" text.
		xmlhttp.Open "GET", BaseURL, False
		'Indicate that the body of the request contains form data
		xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		'Send the data as name/value pairs
		If SendXML(xmlhttp) <> 0 Then
			Set xmlhttp = Nothing
			.Sheets(SheetName).Range("B" & (DayOffset - 6)).Value = "No Response from Environment Canada"
			Exit Sub
		End If
		'Pauses the module while the web data loads.
		While xmlhttp.READYSTATE <> 4
			DoEvents
		Wend
		'Assigns the the website's HTML to the HTML_Data variable.
		HTML_Data = xmlhttp.responseText
		'-----------------------------------------------------------------------------------------------------------------------------'

		''''''''''Extracts the Current Conditions and Watches and Warnings'''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'The irrelevant HTML code is cut.
		'The InStr function searches the code for the string that precedes the relevant data: '<rights>Copyright 2016, Environment Canada</rights>'.
		'The InStr function then returns the number of characters from the start of the HTML code to the start of this string.
		'The Mid function then deletes every character before this number
		If InStr(HTML_Data, "<rights>Copyright 2016, Environment Canada</rights>") > 0 Then
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<rights>Copyright 2016, Environment Canada</rights>"), Len(HTML_Data))
		End If

		'Isolates the watches and warnings string.
		HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<title>") + 7, Len(HTML_Data))
		DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "</title>") - 1)
		'The SheetName variable is recieved from the datepicker in the 'Update' form.
		.Sheets(SheetName).Range("B" & (DayOffset - 6)).Value = DataString

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

			.Sheets(SheetName).Range(Cell(i)).Value = DataString 'Set the value of the appropriate cell
		next i

		'-----------------------------------------------------------------------------------------------------------------------------'

		''''''''''Extracts the Long Term Forecast'''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''

		For i = 1 to 13
			Day = DayOffset + i

			'Isolates the 7 day forecast day.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<title>") + 7, Len(HTML_Data))
			DataString = Mid(HTML_Data, 1, InStr(HTML_Data, ":"))
			.Sheets(SheetName).Range("A" & Day).Value = DataString
			'Isolates the 7 day forecast data.
			'Chr(34) returns a double quotation mark (") and is used to prevent runtime errors.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<summary type=" & Chr(34) & "html" & Chr(34) & ">") + 21, Len(HTML_Data))
			DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "Forecast issued") - 1)
			.Sheets(SheetName).Range("B" & Day).Value = DataString
		next i

		'Isolates the Long Term Forecast time issued.
		HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "Forecast") + 0, Len(HTML_Data))
		DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "</summary>") - 1)
		.Sheets(SheetName).Range("A" & DayOffset).Value = DataString

		'Once the 7th day's forecast is loaded, the xmlhttp is set to 'Nothing' to prevent caching and the module closes.
		Set xmlhttp = Nothing

	End With
End Sub

Sub OttawaScraper(SheetName As String)
	Call ECWeatherScraper(SheetName, "http://weather.gc.ca/rss/city/on-118_e.xml", ECStart)
End Sub