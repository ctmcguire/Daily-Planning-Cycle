Option Explicit

'The sheetdate variable is defined as cell B6 on the sheet that the Update Website button was pressed.
'This is defined in the 'Assign Macro' formula.
Sub Run_WebUpdate(sheetdate As String)
	'-----------------------------------------------------------------------------------------------------------------------------'
	'Please send any questions or feedback to cmcguire@mvc.on.ca
	'-----------------------------------------------------------------------------------------------------------------------------'

	'The Run_WebUpdate function is used to populate a workbook that is used to upload data to an SQL database hosted on mvc.on.ca.
	'The workbook location may have to be adjusted when running on a different computer.
	'1.  Navigate to https://dev.mysql.com/downloads/connector/odbc/ and download the Windows Windows (x86, 32-bit), MSI Installer and Windows (x86, 64-bit), MSI Installer.  You may require both drivers in the event that your 64-bit Windows is running 32-bit Excel.
	'2.  Open the downloaded files and follow the installation prompts.
	'3.  Navigate to C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Administrative Tools or type 'ODBC' into the start menu search and open 'ODBC Data Source Administrator (32-bit)'.
	'4.  Click on 'System DSN', then click on 'Add'.
	'5.  Under 'Data Source Name:' enter a relevant name, no spaces.
	'6.  Select 'TCP/IP Server:' and enter your database's domain name (mvc.on.ca) or IP address.
	'7.  Under 'User:' enter a user that has permission to edit the database (this is not necessarily the same as the user name used to log into your website management portal).
	'To find an eligible user, log into cPanel and click on 'MySQL Databases'.
	'8.  Under 'Password:' enter the password associated with the user.
	'9.  Under 'Database:' type the database name you are trying to access.
	'10. Click 'Test'.  "Connection Successful" should pop up.
	'11. Edit the 'WebUpdate' macro's line 'LevelsConn.ConnectionString =' so that it references your database.
	'12. If the 'WebUpdate' macro fails to connect, navigate to C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Administrative Tools or type 'ODBC' into the start menu search and open 'ODBC Data Source Administrator (64-bit)' and repeat steps 4-10.
	'-----------------------------------------------------------------------------------------------------------------------------'
	''''''''''Textbox Debugger'''''''''''''
	'''''''''''''''''''''''''''''''''''''''
	'Insert these lines anywhere in the code to view strings:
	'Debug_Text.TextBox1 = GaugeName(i)
	'Debug_Text.Show
	'-----------------------------------------------------------------------------------------------------------------------------'

	Call Gauges.InitializeGauges
	'The Status Bar is located on the bottom left corner of the Excel window.  It's default status is 'READY'.
	'The Status Bar Displays 'Processing Request...' until the UpdateDPC subroutine has ended.
	Application.StatusBar = "Processing Request..."
	'Screen Updating is turned off to speed up the processing time.
	Application.ScreenUpdating = False
	'Sheet calculations are turned off to speed up the processing time.
	Application.Calculation = xlCalculationManual

	'The 'i' variable is used to navigate the rows of the workbook.
	Dim i As Integer
	'The 'InputDate' variable references the sheet where the Upload to Website was clicked.
	Dim InputDate As String
	InputDate = Format(sheetdate, "mmm d")

	'-----------------------------------------------------------------------------------------------------------------------------'
	Dim LevelsConn As ADODB.Connection
	Set LevelsConn = New ADODB.Connection
	LevelsConn.ConnectionString = "DRIVER={MySQL ODBC 5.3 Unicode Driver};SERVER=mvc.on.ca;DATABASE=mvconc55_mvclevels;UID=mvconc55_levels1;PWD=4z9!yA;OPTION=3"
	LevelsConn.Open
	'-----------------------------------------------------------------------------------------------------------------------------'

	'The With statement ensures the macro references the daily planning cycle workbook.
	With ThisWorkbook
		For i = 0 To UBound(FlowGauges)
			'If the value exists, the Date, Time, Value and Historical average are uploaded.
			If .Sheets("Raw2").Range("E" & (flowStart + i)) < .Sheets(InputDate).Range("E" & (flowStart + i)) Then _
				Call Run_SQL(i + flowStart, InputDate, FlowGauges(i).Name, LevelsConn)
		Next i

		For i = 0 To UBound(DailyGauges)
			If .Sheets("Raw2").Range("E" & (dailyStart + i)) < .Sheets(InputDate).Range("E" & (dailyStart + i)) Then _
				Call Run_SQL(i + dailyStart, InputDate, DailyGauges(i).Name, LevelsConn)
		Next i

		For i = 0 To UBound(WeeklyGauges)
			'These values are the same between the Raw2 sheet and the other sheets, so this if statement instead checks if the value is positive
			If 0 < .Sheets(InputDate).Range("E" & (weeklyStart + i)) Then _
				Call Run_SQL(i + weeklyStart, InputDate, WeeklyGauges(i).Name, LevelsConn) 'NoRain is set to true because there is no precipitation measurement for the weekly values
		Next i

		LevelsConn.Close

		MsgBox "The requested data has been uploaded to the website. Please visit: http://mvc.on.ca/water-levels-app/levels-table-option/ataglance.php to ensure accuracy."

		'The previously adjusted modes are returned to their default state.
		Application.StatusBar = False
		Application.Calculation = xlCalculationAutomatic
		Application.ScreenUpdating = True

	End With
End Sub

'/* 
' * This Subroutine creates and executes the SQL query for one row, based on the values it is given
' * 
' * @param i			- The row of the data being transferred to the SQL database
' * @param InputDate	- The date (and sheetname) of the worksheet whose data is being transferred to the SQL database
' * @param GaugeName	- The name of the gauge that corresponds to the i-value
' * @param LevelsConn	- A connection to the SQL database that can be used to run the SQL queries
' */
Sub Run_SQL(i As Integer, InputDate As String, GaugeName As String, LevelsConn As ADODB.Connection)
	Dim strSQL As String 'String to store the SQL query
	Dim havg As String
	Dim Rain As String

	With ThisWorkbook
		havg = "NULL" 'Set the historical average to NULL by default
		Rain = "NULL" 'Precip is NULL by default

		'If there is data for the historical average, update the value of havg
		If Not IsEmpty(.Sheets(InputDate).Range("I" & i)) Then _
			havg = esc(.Sheets(InputDate).Range("I" & i))
		'If there is data for the precipitation, update the value of Rain
		If Not (IsEmpty(ThisWorkbook.Sheets(InputDate).Range("K" & i)) Or GaugeName = "Mississippi Lake" Or GaugeName = "Kashwakamak Lake" Or 0 < InStr(GaugeName, "(weekly)")) Then _
			Rain = esc(.Sheets(InputDate).Range("K" & i))

		strSQL = "INSERT INTO mvconc55_mvclevels.data (id, gauge, date, time, datainfo, historicalaverage, precipitation) " & _
			"VALUES (NULL, '" & esc(GaugeName) & "', '" & esc(Format(.Sheets(InputDate).Range("B" & i), "yyyy-mm-dd")) & "', '" & _
			esc(Format(.Sheets(InputDate).Range("C" & i), "h:mm AMPM")) & "', " & esc(.Sheets(InputDate).Range("E" & i)) & ", " & _
			havg & ", " & Rain & ") ON DUPLICATE KEY UPDATE time='" & _
			esc(Format(.Sheets(InputDate).Range("C" & i), "h:mm AMPM")) & "', datainfo=" & _
			esc(.Sheets(InputDate).Range("E" & i)) & ", historicalaverage=" & havg & ", precipitation=" & Rain & ";"
		LevelsConn.Execute strSQL
	End With
End Sub

Function esc(txt As String)
	esc = Trim(Replace(txt, "'", "\'"))
End Function