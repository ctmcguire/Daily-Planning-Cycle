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

Private Sub AccuWeatherScraper(SheetName As String, BaseUrl As String, DayOffset As Integer)
	'-----------------------------------------------------------------------------------------------------------------------------'
	'Please send any questions or feedback to cmcguire@mvc.on.ca
	'-----------------------------------------------------------------------------------------------------------------------------'
	''''''''''Textbox Debugger'''''''''''''
	'''''''''''''''''''''''''''''''''''''''
	'Insert these lines anywhere in the code to view the imported HTML data:
	'Debug_Text.TextBox1 = HTML_Data
	'Debug_Text.Show
	'-----------------------------------------------------------------------------------------------------------------------------'
	'The AccuWeatherScraper module loads the 5 day forecast data from accuweather.com.
	'-----------------------------------------------------------------------------------------------------------------------------'
	'the xmlhttp object interacts with the web server to retrieve the data
	Dim xmlhttp As Object
	'The HTML_Data variable stores all the response HTML text from the website as a string.
	Dim HTML_Data As String
	'The DataString variable is the parsed HTML code that is inserted into the Daily Planning Cycle
	Dim DataString As String
	'The Day variable is used to navigate the rows.
	Dim Day As Integer
	'The Temp variable is used to store the temperature data
	Dim Temp As String
	'The URL variable reduces the chance of error when cycling through the precipitation pages.
	Dim URL As String
	'The Day/Night variables are used to calculate total 24-hour precipitation, rain and snow.
	Dim DayPrecip As Double
	Dim DayRain As Double
	Dim DaySnow As Double
	Dim NightPrecip As Double
	Dim NightRain As Double
	Dim NightSnow As Double

	Dim Deg As Integer 'Variable used to store the index of the degrees symbol when parsing the html response
	'The i variable is used as a counter in the loop that grabs the precipitation values.
	Dim i As Integer

	Day = DayOffset
	'The With statement is used to ensure the macro does not modify other workbooks that may be open.
	With ThisWorkbook
		If .Sheets(SheetName).Range("B" & Day).Value <> "No Response from AccuWeather" And .Sheets(SheetName).Range("B" & Day).Value <> "" Then _
			Goto Forecast5Day
		'-----------------------------------------------------------------------------------------------------------------------------'
		''''''''''Loads the web data into VBA'''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''
		'Creates the xmlhttp object that interacts with the website. .ServerXMLHTTP60 is used so the page data is not cached.
		Set xmlhttp = New MSXML2.ServerXMLHTTP60
		'Indicates that page that will receive the request and the type of request being submitted.  Your location's link can be found by searching for your location at 'accuweather.com' and clicking 'Extended'.
		xmlhttp.Open "GET", BaseUrl, False
		'Indicate that the body of the request contains form data
		xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		'Send the data as name/value pairs
		If SendXML(xmlhttp) <> 0 Then
			Set xmlhttp = Nothing
			.Sheets(SheetName).Range("B" & Day).Value = "No Response from AccuWeather"
			Goto Forecast5Day
		End If
		'Pauses the module while the web data loads.
		While xmlhttp.READYSTATE <> 4
			DoEvents
		Wend
		'Assigns the the website's HTML to the HTML_Data variable.
		HTML_Data = xmlhttp.responseText

		'-----------------------------------------------------------------------------------------------------------------------------'
		''''''''''Extracts the Forecast Location''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''
		'Isolates the forecast location
		'The InStr function searches the code for the string that precedes the current conditions observation time: 'timestampApp_local'.
		'The InStr function then returns the number of characters from the start of the HTML code to the start of this string.
		'The Mid function then deletes every character before this number.
		'Chr(34) returns a double quotation mark (") and is used to prevent runtime errors.
		HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "class=" & Chr(34) & "current-city" & Chr(34) & "><h1>") + 25, Len(HTML_Data))
		DataString = Mid(HTML_Data, 1, InStr(HTML_Data, ",") - 1) + " 2016 AccuWeather, Inc. All Rights Reserved."
		'The SheetName variable is recieved from the datepicker in the 'Update' form
		.Sheets(SheetName).Range("A" & Day).Value = DataString

		'Cuts the extra HTML code.
		If InStr(HTML_Data, "<!-- /.feed-controls -->") > 0 Then
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<!-- /.feed-controls -->"), Len(HTML_Data))
		End If

		'-----------------------------------------------------------------------------------------------------------------------------'
		''''''''''Extracts the 5 Day Forecasted Highs and Lows'''''''''''''
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'This for loop ensures the entire 5 day forecast is extracted before proceeding.
		For i = 1 to 5
			Day = DayOffset + i

			'Isolates the forecast date.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<a href=" & Chr(34) & "#" & Chr(34) & ">") + 12, Len(HTML_Data))
			DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "</a></h3>") - 1)
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<h4>") + 4, Len(HTML_Data))
			DataString = DataString + " " + Mid(HTML_Data, 1, InStr(HTML_Data, "</h4>") - 1)
			.Sheets(SheetName).Range("A" & Day).Value = DataString

			'Isolates the forecast high.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<span class=" & Chr(34) & "large-temp" & Chr(34) & ">") + 25, Len(HTML_Data))

			Temp = Mid(HTML_Data, 1, InStr(HTML_Data, "&deg") - 1)
			.Sheets(SheetName).Range("C" & Day).Value = Temp

			'Isolates the forecast low.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<span class=" & Chr(34) & "small-temp" & Chr(34) & ">") + 26, Len(HTML_Data))
			Temp = Mid(HTML_Data, 1, InStr(HTML_Data, "&deg") - 1)
			.Sheets(SheetName).Range("D" & Day).Value = Temp

			'Isolates the forecast condition.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<span class=" & Chr(34) & "cond" & Chr(34) & ">") + 19, Len(HTML_Data))
			DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "</span>") - 1)
			.Sheets(SheetName).Range("B" & Day).Value = DataString
		next i
		
		'-----------------------------------------------------------------------------------------------------------------------------'
		Day = DayOffset

		'Adds a VBA time stamp to the weather since time is not published on the webpage.
		.Sheets(SheetName).Range("B" & Day).Value = Format(DateTime.Now, "yyyy-MM-d hh:mm:ss")

		Set xmlhttp = Nothing 'Clears the object from memory

		Forecast5Day:
		'-----------------------------------------------------------------------------------------------------------------------------'
		''''''''''Extracts the 5 day forecasted data from separate pages'''''''''''''
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		'The i variable navigates to the corresponding forecast day.
		For i = 1 to 5
			Day = DayOffset + i

			'Your Location's link can be found by searching for your location at 'accuweather.com' and clicking 'Extended'.
			URL = BaseUrl & "?day=" & i

			If .Sheets(SheetName).Range("E" & Day).Value <> "No Response from AccuWeather" And .Sheets(SheetName).Range("E" & Day).Value <> "" Then _
				Goto Continue

			'Creates the xmlhttp object that interacts with the website. .ServerXMLHTTP60 is used so the page data is not cached.
			Set xmlhttp = New MSXML2.ServerXMLHTTP60
			With xmlhttp
				.Open "GET", URL, False
				.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
				If SendXML(xmlhttp) <> 0 Then
					Set xmlhttp = Nothing
					ThisWorkbook.Sheets(SheetName).Range("E" & Day).Value = "No Response from AccuWeather"
					Goto Continue
				End If
				While .READYSTATE <> 4
					DoEvents
				Wend
				HTML_Data = .responseText
			End With

			'Cuts the HTML code to the precipitation
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<!-- /.feed-tabs -->"), Len(HTML_Data))

			'Isolates the RealFeal temperature.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "RealFeel&#174;") + 14, Len(HTML_Data))
			Deg = InStr(HTML_Data, "&#176;</span>")
			If Deg = 0 Then _
				Deg = InStr(HTML_Data, "&deg;</span>")
			IF Deg = 0 Then
				MsgBox "Error parsing HTML string: degree character not found.  Attempting to parse without it..."
				Deg = InStr(HTML_Data, "</span>")
			End If
			Temp = Mid(HTML_Data, 1, Deg - 1)
			.Sheets(SheetName).Range("E" & Day).Value = Temp

			'Isolates the POP.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, ">Precipitation") + 14, Len(HTML_Data))
			DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "</span>") - 1)
			.Sheets(SheetName).Range("H" & Day).Value = DataString

			'Isolates the Wind.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<strong>") + 8, Len(HTML_Data))
			DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "</strong>") - 1)
			.Sheets(SheetName).Range("F" & Day).Value = DataString

			'Isolates the Wind Gusts.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "Gusts:<strong style=") + 24, Len(HTML_Data))
			DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "</strong>") - 1)
			.Sheets(SheetName).Range("G" & Day).Value = DataString

			'Isolates the day precipitation.
			'This number is added to the night precipitation and the 24-hour total precipitation is added to the DPC.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<li>Precipitation: <strong>") + 27, Len(HTML_Data))
			DayPrecip = Val(Mid(HTML_Data, 1, InStr(HTML_Data, " mm")))

			'Isolates the day rain
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<li>Rain: <strong style=" & Chr(34) & Chr(34) & ">") + 27, Len(HTML_Data))
			DayRain = Val(Mid(HTML_Data, 1, InStr(HTML_Data, " mm")))

			'Isolates the day snow
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<li>Snow: <strong style=" & Chr(34) & Chr(34) & ">") + 27, Len(HTML_Data))
			DaySnow = Mid(HTML_Data, 1, InStr(HTML_Data, " CM") - 1)

			'Isolates the day ice
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<li>Ice: <strong style=" & Chr(34) & Chr(34) & ">") + 26, Len(HTML_Data))
			DayRain = DayRain + Val(Mid(HTML_Data, 1, InStr(HTML_Data, " mm")))

			'Isolates the Night precip
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<li>Precipitation: <strong>") + 27, Len(HTML_Data))
			NightPrecip = DayPrecip + Val(Mid(HTML_Data, 1, InStr(HTML_Data, " mm")))
			.Sheets(SheetName).Range("K" & Day).Value = NightPrecip

			'Isolates the Night rain
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<li>Rain: <strong style=" & Chr(34) & Chr(34) & ">") + 27, Len(HTML_Data))
			NightRain = DayRain + Val(Mid(HTML_Data, 1, InStr(HTML_Data, " mm")))

			'Isolates the Night snow
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<li>Snow: <strong style=" & Chr(34) & Chr(34) & ">") + 27, Len(HTML_Data))
			NightSnow = DaySnow + Val(Mid(HTML_Data, 1, InStr(HTML_Data, " CM")))
			.Sheets(SheetName).Range("J" & Day).Value = NightSnow

			'Isolates the Night ice
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<li>Ice: <strong style=" & Chr(34) & Chr(34) & ">") + 26, Len(HTML_Data))
			NightRain = NightRain + Val(Mid(HTML_Data, 1, InStr(HTML_Data, " mm")))
			.Sheets(SheetName).Range("I" & Day).Value = NightRain

			Set xmlhttp = Nothing
			Continue: 'Label so Goto lines in error checking can go here
		next i
	End With
End Sub

Sub GeneralScraper(SheetName As String, LocationURL As String, Optional RowNo As Integer = 0)
	If RowNo = 0 Then _
		RowNo = NextWeather
	Call AccuWeatherScraper(SheetName, "http://www.accuweather.com/en/ca/" & LocationURL, RowNo)
	NextWeather = RowNo + AccuCount + 2
End Sub

Sub CPScraper(SheetName As String)
	Call AccuWeatherScraper(SheetName, "http://www.accuweather.com/en/ca/carleton-place/k7c/daily-weather-forecast/55438", AccuStart)
End Sub


Sub CloyneScraper(SheetName As String)
	Call AccuWeatherScraper(SheetName, "http://www.accuweather.com/en/ca/cloyne/k0h/daily-weather-forecast/2291535", CloyneAccuStart)
End Sub