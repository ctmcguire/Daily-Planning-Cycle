Option Explicit

Sub AccuWeatherScraper(SheetName As String)
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
'The i variable is used as a counter in the loop that grabs the precipitation values.
Dim i As Integer

Day = 81
'The With statement is used to ensure the macro does not modify other workbooks that may be open.
With ThisWorkbook
'-----------------------------------------------------------------------------------------------------------------------------'

''''''''''Loads the web data into VBA'''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''
'Creates the xmlhttp object that interacts with the website. .ServerXMLHTTP60 is used so the page data is not cached.
Set xmlhttp = New MSXML2.ServerXMLHTTP60
'Indicates that page that will receive the request and the type of request being submitted.  Your location's link can be found by searching for your location at 'accuweather.com' and clicking 'Extended'.
xmlhttp.Open "GET", "http://www.accuweather.com/en/ca/carleton-place/k7c/daily-weather-forecast/55438", False
'Indicate that the body of the request contains form data
xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
'Send the data as name/value pairs
xmlhttp.send
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
Day = Day + 1

'Cuts the extra HTML code.
If InStr(HTML_Data, "<!-- /.feed-controls -->") > 0 Then
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<!-- /.feed-controls -->"), Len(HTML_Data))
End If

'-----------------------------------------------------------------------------------------------------------------------------'

''''''''''Extracts the 5 Day Forecasted Highs and Lows'''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
next_temp:
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

Day = Day + 1
'This if loop ensures the entire 5 day forecast is extracted before proceeding.
If InStr(HTML_Data, "<span class=" & Chr(34) & "small-temp" & Chr(34) & ">") > 0 Then GoTo next_temp:

'-----------------------------------------------------------------------------------------------------------------------------'

Day = 81

'Adds a VBA time stamp to the weather since time is not published on the webpage.
.Sheets(SheetName).Range("B" & Day).Value = Format(DateTime.Now, "yyyy-MM-d hh:mm:ss")
Day = Day + 1

'-----------------------------------------------------------------------------------------------------------------------------'

''''''''''Extracts the 5 day forecasted data from separate pages'''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'The i variable navigates to the corresponding forecast day.
i = 1

next_row:
'Isolates the RealFeal temperature.
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "RealFeel&#174;") + 14, Len(HTML_Data))
Temp = Mid(HTML_Data, 1, InStr(HTML_Data, "&#176;</span>") - 1)
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

Day = Day + 1
i = i + 1
Set xmlhttp = Nothing

'Once the 5 day forecast data has been extracted, the module closes.
If i > 5 Then Exit Sub

'Your Location's link can be found by searching for your location at 'accuweather.com' and clicking 'Extended'.
URL = "http://www.accuweather.com/en/ca/carleton-place/k7c/daily-weather-forecast/55438?day=" & i

'Creates the xmlhttp object that interacts with the website. .ServerXMLHTTP60 is used so the page data is not cached.
Set xmlhttp = New MSXML2.ServerXMLHTTP60
With xmlhttp
    .Open "GET", URL, False
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .send
    While .READYSTATE <> 4
        DoEvents
    Wend
    HTML_Data = .responseText
End With

'Cuts the HTML code to the precipitation
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<!-- /.feed-tabs -->"), Len(HTML_Data))

'If InStr(HTML_Data, "<strong class=" & Chr(34) & "temp" & Chr(34) & ">") > 0 Then
GoTo next_row:

End With
End Sub
