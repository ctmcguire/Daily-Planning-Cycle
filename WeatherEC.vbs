Option Explicit

Sub ECWeatherScraper(SheetName As String)
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
'The With statement is used to ensure the macro does not modify other workbooks that may be open.
With ThisWorkbook
Day = 105
'-----------------------------------------------------------------------------------------------------------------------------'

''''''''''Loads the web data into VBA'''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''
'Creates the xmlhttp object that interacts with the website. .ServerXMLHTTP60 is used so the page data is not cached.
'Loads the weather page and saves the HTML data as the variable HTML_Data
Set xmlhttp = New MSXML2.ServerXMLHTTP60
'Indicates that page that will receive the request and the type of request being submitted.
'Your location's link can be found by searching for a local forecast at: http://weather.gc.ca/canada_e.html
'After the local forecast has loaded, click on the RSS Weather link underneath the historical data and adjacent to the 'Follow:" text.
xmlhttp.Open "GET", "http://weather.gc.ca/rss/city/on-118_e.xml", False
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
.Sheets(SheetName).Range("B" & Day).Value = DataString
Day = Day + 1

next_row:
'Isolates the observation and current condition data.
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "</b>") + 5, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "<br/>") - 1)
.Sheets(SheetName).Range("B" & Day).Value = DataString
Day = Day + 1
'This 'If' loop repeates the code from the 'next_row:' to extract the observation and current condition data.
If InStr(HTML_Data, "Condition") > 0 Then GoTo next_row:

'Isolates the current temperature data.
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<b>Temperature:</b>") + 20, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "<br/>") - 8)
.Sheets(SheetName).Range("B" & Day).Value = DataString

next_string:
'Isolates the Pressure, Visibility and Humidity data
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "</b>") + 5, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "<br/>") - 1)
.Sheets(SheetName).Range("E" & Day).Value = DataString
Day = Day + 1

If InStr(HTML_Data, "<b>Humidity:</b>") > 0 Then
GoTo next_string:

'This If statement extracts the Wind Chill data when it exists
ElseIf InStr(HTML_Data, "<b>Wind Chill:</b>") > 0 Then
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<b>Wind Chill:</b>") + 19, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "<br/>") - 1)
.Sheets(SheetName).Range("B" & Day - 2).Value = DataString
'Day = Day + 1

Else
.Sheets(SheetName).Range("B" & Day - 2).Value = "N/A"
'Day = Day + 1
End If

Day = Day - 1

'Isolates the Dewpoint data
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<b>Dewpoint:</b>") + 17, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "<br/>") - 8)
.Sheets(SheetName).Range("H" & Day - 2).Value = DataString
'Day = Day + 1

'Isolates the Wind data
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<b>Wind:</b>") + 13, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "<br/>") - 1)
.Sheets(SheetName).Range("H" & Day - 1).Value = DataString
'Day = Day + 1

'Isolates the Air Quality data
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "</b>") + 5, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "<br/>") - 2)
.Sheets(SheetName).Range("B" & Day).Value = DataString

Day = Day + 2

'-----------------------------------------------------------------------------------------------------------------------------'

''''''''''Extracts the Long Term Forecast'''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
nextLT_row:
'Isolates the 7 day forecast day.
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<title>") + 7, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, ":") - 0)
.Sheets(SheetName).Range("A" & Day).Value = DataString
'Isolates the 7 day forecast data.
'Chr(34) returns a double quotation mark (") and is used to prevent runtime errors.
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "<summary type=" & Chr(34) & "html" & Chr(34) & ">") + 21, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "Forecast issued") - 1)
.Sheets(SheetName).Range("B" & Day).Value = DataString
Day = Day + 1
'This 'If' loop repeates the code from the 'next_LTrow:' until the entire long term forecast is extracted.
If InStr(HTML_Data, "<summary type=" & Chr(34) & "html" & Chr(34) & ">") > 0 Then GoTo nextLT_row:

'Isolates the Long Term Forecast time issued.
HTML_Data = Mid(HTML_Data, InStr(HTML_Data, "Forecast") + 0, Len(HTML_Data))
DataString = Mid(HTML_Data, 1, InStr(HTML_Data, "</summary>") - 1)
.Sheets(SheetName).Range("A" & Day - 14).Value = DataString

'Once the 7th day's forecast is loaded, the xmlhttp is set to 'Nothing' to prevent caching and the module closes.
Set xmlhttp = Nothing

End With
End Sub
