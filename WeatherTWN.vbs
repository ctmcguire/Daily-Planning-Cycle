Option Explicit

Sub TWNWeatherScraper(SheetName As String)
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
'The With statement is used to ensure the macro does not modify other workbooks that may be open.
With ThisWorkbook
Day = 88
'-----------------------------------------------------------------------------------------------------------------------------'

''''''''''Loads the web data into VBA'''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''
'Creates the xmlhttp object that interacts with the website. .ServerXMLHTTP60 is used so the page data is not cached.
Set xmlhttp = New MSXML2.ServerXMLHTTP60
'Indicates that page that will receive the request and the type of request being submitted.  Your location's link can be found at: http://legacyweb.theweathernetwork.com/rss/
xmlhttp.Open "GET", "http://legacyweb.theweathernetwork.com/dataaccess/citypage/json/caon0119", False
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
.Sheets(SheetName).Range("C" & Day).Value = DPCTimeStamp
Day = Day + 1

'Isolates the Current Temperature
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "temperature_c") + 15, Len(HTML_Data))
'Chr(34) returns double quotation marks (") and is used to prevent runtime errors.
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1)
.Sheets(SheetName).Range("B" & Day).Value = DataString
Day = Day + 1

'Isolates the 'Feels like' Temperature
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "feelsLike_c") + 13, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1)
.Sheets(SheetName).Range("B" & Day).Value = DataString
Day = Day - 1

'Isolates the Wind Direction
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "windDirection") + 16, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, Chr(34) & ",") - 1)
'Translates the wind direction to english
DataString = Replace(DataString, "O", "W")
'Isolates the Wind Speed
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "windSpeed_kmh") + 15, Len(HTML_Data))
'Merges the wind direction and speed into one string.
DataString = DataString + " " + Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1) + " km/h"
.Sheets(SheetName).Range("E" & Day).Value = DataString
Day = Day + 1

'Isolates the Wind gusts
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "windGustSpeed_kmh") + 19, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1)
.Sheets(SheetName).Range("E" & Day).Value = DataString + " km/h"

'Isolates the current conditions location
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "name_en") + 10, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, Chr(34) & ",") - 1)
.Sheets(SheetName).Range("B" & Day - 2).Value = DataString

Day = Day + 1
'-----------------------------------------------------------------------------------------------------------------------------'

''''''''''Extracts the Short Term Forecast'''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Isolates the Short Term Forecast time
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "timestamp_local") + 18, Len(HTML_Data))
WebTimeStamp = Mid(HTML_Data, 1, InStr(HTML_Data, "tzbias") - 3)
DPCTimeStamp = (WebTimeStamp / (86400000) + 25569)
.Sheets(SheetName).Range("B" & Day).Value = DPCTimeStamp
Day = Day + 1

next_STrow:
'Isolates the Forecast date
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "timestampApp_local") + 21, Len(HTML_Data))
WebTimeStamp = Mid(HTML_Data, 1, InStr(HTML_Data, "icon") - 3)
DPCTimeStamp = (WebTimeStamp / (86400000) + 25569)
.Sheets(SheetName).Range("A" & Day).Value = DPCTimeStamp

'Isolates the Short Term Forecast Temperature
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "temperature_c") + 15, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1)
.Sheets(SheetName).Range("C" & Day).Value = DataString

'Isolates the 'Feels like' Temperature
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "feelsLike_c") + 13, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1)
.Sheets(SheetName).Range("D" & Day).Value = DataString

'Isolates the POP
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "pop_percent") + 13, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1) + "%"
.Sheets(SheetName).Range("E" & Day).Value = DataString

'Isolates the Wind Direction
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "windDirection") + 16, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, Chr(34) & ",") - 1)
'Translates the wind direction to english
DataString = Replace(DataString, "O", "W")

'Isolates the Wind Speed
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "windSpeed_kmh") + 15, Len(HTML_Data))
DataString = DataString + " " + Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1) + " km/h"
.Sheets(SheetName).Range("F" & Day).Value = DataString

'Isolates the Wind gusts
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "windGustSpeed_kmh") + 19, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1) + " km/h"
.Sheets(SheetName).Range("G" & Day).Value = DataString

'Isolates the Forecasted Rain
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, Chr(34) & "rain" & Chr(34)) + 7, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1)
.Sheets(SheetName).Range("H" & Day).Value = DataString

'Isolates the Forecasted Snow
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, Chr(34) & "snow" & Chr(34)) + 7, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1)
.Sheets(SheetName).Range("I" & Day).Value = DataString

Day = Day + 1

'This 'If' loop repeates the code from the 'next_STrow:' until the entire short term forecast is extracted.
If InStr(HTML_Data, "temperature_c") > 0 Then GoTo next_STrow:

'-----------------------------------------------------------------------------------------------------------------------------'

''''''''''Extracts the Long Term Forecast'''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Isolates the Long Term Forecast time issued
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "timestamp_local") + 18, Len(HTML_Data))
WebTimeStamp = Mid(HTML_Data, 1, InStr(HTML_Data, "tzbias") - 3)
DPCTimeStamp = (WebTimeStamp / (86400000) + 25569)
.Sheets(SheetName).Range("B" & Day).Value = DPCTimeStamp
Day = Day + 1

next_LTrow:
'Isolates the Long Term Forecast date
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "timestampApp_local") + 21, Len(HTML_Data))
WebTimeStamp = Mid(HTML_Data, 1, InStr(HTML_Data, "icon") - 3)
DPCTimeStamp = (WebTimeStamp / (86400000) + 25569)
.Sheets(SheetName).Range("A" & Day).Value = DPCTimeStamp

'Isolates the Long Term Forecast Low
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "temperatureMin_c") + 18, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1)
.Sheets(SheetName).Range("D" & Day).Value = DataString

'Isolates the Long Term Forecast High
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "temperatureMax_c") + 18, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1)
.Sheets(SheetName).Range("C" & Day).Value = DataString

'Isolates the 'Feels like' Temperature
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "feelsLike_c") + 13, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1)
.Sheets(SheetName).Range("E" & Day).Value = DataString

'Isolates the Wind Direction
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "windDirection") + 16, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, Chr(34) & ",") - 1)
'Translates the wind direction to english
DataString = Replace(DataString, "O", "W")
'Isolates the Wind Speed
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "windSpeed_kmh") + 15, Len(HTML_Data))
DataString = DataString + " " + Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1) + " km/h"
.Sheets(SheetName).Range("F" & Day).Value = DataString

'Isolates the Day POP
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "popDay_percent") + 16, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1) + "%"
.Sheets(SheetName).Range("G" & Day).Value = DataString

'Isolates the Rain
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, Chr(34) & "rain" & Chr(34)) + 7, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1)
.Sheets(SheetName).Range("H" & Day).Value = DataString

'Isolates the Snow
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, Chr(34) & "snow" & Chr(34)) + 7, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "," & Chr(34)) - 1)
.Sheets(SheetName).Range("I" & Day).Value = DataString

Day = Day + 1

If InStr(HTML_Data, "temperatureMin_c") > 0 Then GoTo next_LTrow:

'Once the 7th day's forecast is loaded, the xmlhttp is set to 'Nothing' to prevent caching and the module closes.
Set xmlhttp = Nothing

End With
End Sub
