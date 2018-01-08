Option Explicit

Private Function SendXML(xmlhttp As Object) As Boolean
	On Error GoTo OnError
	SendXML = False
	With xmlhttp
		.send
		'SendXML = .waitForResponse(60000) 'This line either sets SendXML to True, sets it to False, or gets skipped, which leaves SendXML as False
	End With
	SendXML = True
	OnError:
End Function

Private Sub AccuWeatherScraper(SheetName As String, BaseUrl As String, DayOffset As Integer, Optional IsAuto As Boolean = False)
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
	'The Url variable reduces the chance of error when cycling through the precipitation pages.
	Dim Url As String
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
		'Adds a VBA time stamp to the weather since time is not published on the webpage.
		If IsEmpty(.Sheets(SheetName).Range("B" & Day).Value) Then _
			.Sheets(SheetName).Range("B" & Day).Value = Format(Now, "yyyy-MM-d hh:mm:ss")

		'-----------------------------------------------------------------------------------------------------------------------------'
		''''''''''Extracts the 5 day forecasted data from separate pages'''''''''''''
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		Call DebugLogging.PrintMsg("AccuWeather - Getting forecast for 5 days...")

		'The i variable navigates to the corresponding forecast day.
		For i = 1 to 5
			Day = DayOffset + i

			'Your Location's link can be found by searching for your location at 'accuweather.com' and clicking 'Extended'.
			Url = BaseUrl & "?day=" & i

			If .Sheets(SheetName).Range("A" & Day).Value <> "No Response from AccuWeather" And .Sheets(SheetName).Range("A" & Day).Value <> "" Then _
				Goto Continue

			'Creates the xmlhttp object that interacts with the website. .ServerXMLHTTP60 is used so the page data is not cached.
			Set xmlhttp = New MSXML2.XMLHTTP60
			With xmlhttp
				.Open "GET", Url, False
				.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
				If Not SendXML(xmlhttp) Then
					Set xmlhttp = Nothing
					ThisWorkbook.Sheets(SheetName).Range("A" & Day).Value = "No Response from AccuWeather"
					Goto Continue
				End If
				HTML_Data = .responseText
			End With

			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "class=" & Chr(34) & "current-city" & Chr(34) & "><h1>") + 25, Len(HTML_Data))
			DataString = Mid(HTML_Data, 1, InStr(HTML_Data, ",") - 1) + " 2016 AccuWeather, Inc. All Rights Reserved."
			'The SheetName variable is recieved from the datepicker in the 'Update' form
			.Sheets(SheetName).Range("A" & DayOffset).Value = DataString

			If InStr(HTML_Data, "<!-- /.feed-controls -->") > 0 Then
				HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<!-- /.feed-controls -->"), Len(HTML_Data))
			End If
			'Isolates the forecast date.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, Url), Len(HTML_Data))
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<a href=" & Chr(34) & "#" & Chr(34) & ">") + 12, Len(HTML_Data))
			DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "</a></h3>") - 1)
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<h4>") + 4, Len(HTML_Data))
			DataString = DataString + " " + Mid(HTML_Data, 1, InStr(HTML_Data, "</h4>") - 1)
			.Sheets(SheetName).Range("A" & Day).Value = DataString

			'Isolates the forecast high.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<span class=" & Chr(34) & "large-temp" & Chr(34) & ">") + 25, Len(HTML_Data))

			Temp = Mid(HTML_Data, 1, InStr(HTML_Data, "&deg") - 1)
			If IsEmpty(.Sheets(SheetName).Range("C" & Day)) Then _
				.Sheets(SheetName).Range("C" & Day).Value = Temp

			'Isolates the forecast low.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<span class=" & Chr(34) & "small-temp" & Chr(34) & ">") + 26, Len(HTML_Data))
			Temp = Mid(HTML_Data, 1, InStr(HTML_Data, "&deg") - 1)
			If IsEmpty(.Sheets(SheetName).Range("D" & Day)) Then _
				.Sheets(SheetName).Range("D" & Day).Value = Temp

			'Isolates the forecast condition.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<span class=" & Chr(34) & "cond" & Chr(34) & ">") + 19, Len(HTML_Data))
			DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "</span>") - 1)
			If IsEmpty(.Sheets(SheetName).Range("B" & Day)) Then _
				.Sheets(SheetName).Range("B" & Day).Value = DataString

			'Cuts the HTML code to the precipitation
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<!-- /.feed-tabs -->"), Len(HTML_Data))

			'Isolates the RealFeal temperature.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "RealFeel&#174;") + 14, Len(HTML_Data))
			Deg = InStr(HTML_Data, "&#176;</span>")
			If Deg = 0 Then _
				Deg = InStr(HTML_Data, "&deg;</span>")
			IF Deg = 0 Then
				If Not IsAuto Then _
					MsgBox "Error parsing HTML string: degree character not found.  Attempting to parse without it..."
				Call DebugLogging.PrintMsg("AccuWeather - Error parsing HTML string: degree character not found.  Attempting to parse without it...")
				Deg = InStr(HTML_Data, "</span>")
			End If
			Temp = Mid(HTML_Data, 1, Deg - 1)
			If IsEmpty(.Sheets(SheetName).Range("E" & Day)) Then _
				.Sheets(SheetName).Range("E" & Day).Value = Temp

			'Isolates the POP.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, ">Precipitation") + 14, Len(HTML_Data))
			DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "</span>") - 1)
			If IsEmpty(.Sheets(SheetName).Range("H" & Day)) Then _
				.Sheets(SheetName).Range("H" & Day).Value = DataString

			'Isolates the Wind.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<strong>") + 8, Len(HTML_Data))
			DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "</strong>") - 1)
			If IsEmpty(.Sheets(SheetName).Range("F" & Day)) Then _
				.Sheets(SheetName).Range("F" & Day).Value = DataString

			'Isolates the Wind Gusts.
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "Gusts:<strong style=") + 24, Len(HTML_Data))
			DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "</strong>") - 1)
			If IsEmpty(.Sheets(SheetName).Range("G" & Day)) Then _
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
			If IsEmpty(.Sheets(SheetName).Range("K" & Day)) Then _
				.Sheets(SheetName).Range("K" & Day).Value = NightPrecip

			'Isolates the Night rain
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<li>Rain: <strong style=" & Chr(34) & Chr(34) & ">") + 27, Len(HTML_Data))
			NightRain = DayRain + Val(Mid(HTML_Data, 1, InStr(HTML_Data, " mm")))

			'Isolates the Night snow
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<li>Snow: <strong style=" & Chr(34) & Chr(34) & ">") + 27, Len(HTML_Data))
			NightSnow = DaySnow + Val(Mid(HTML_Data, 1, InStr(HTML_Data, " CM")))
			If IsEmpty(.Sheets(SheetName).Range("J" & Day)) Then _
				.Sheets(SheetName).Range("J" & Day).Value = NightSnow

			'Isolates the Night ice
			HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<li>Ice: <strong style=" & Chr(34) & Chr(34) & ">") + 26, Len(HTML_Data))
			NightRain = NightRain + Val(Mid(HTML_Data, 1, InStr(HTML_Data, " mm")))
			If IsEmpty(.Sheets(SheetName).Range("I" & Day)) Then _
				.Sheets(SheetName).Range("I" & Day).Value = NightRain

			Set xmlhttp = Nothing
			Continue: 'Label so Goto lines in error checking can go here
		next i
	End With

	Call DebugLogging.PrintMsg("AccuWeather - 5 day forecasts retrieved.")

End Sub

Sub GeneralScraper(SheetName As String, LocationUrl As String, Optional IsAuto As Boolean = False, Optional RowNo As Integer = 0)
	If RowNo = 0 Then _
		RowNo = NextWeather
	Call AccuWeatherScraper(SheetName, "https://www.accuweather.com/en/ca/" & LocationUrl, RowNo, IsAuto)
	NextWeather = RowNo + AccuCount + 2
End Sub

Sub CPScraper(SheetName As String, Optional IsAuto As Boolean = False)
	Call GeneralScraper(SheetName, "carleton-place/k7c/daily-weather-forecast/55438", IsAuto, AccuStart)
End Sub


Sub CloyneScraper(SheetName As String, Optional IsAuto As Boolean = False)
	Call GeneralScraper(SheetName, "cloyne/k0h/daily-weather-forecast/2291535", IsAuto, CloyneAccuStart)
End Sub