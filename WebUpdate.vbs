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
	Const FlowOffset As Integer = 6
	Const FlowGaugeCount As Integer = 12
	Dim FlowGaugeName(FlowGaugeCount) As String
	FlowGaugeName(0) = "Myers Cave flow"
	FlowGaugeName(1) = "Buckshot Creek flow"
	FlowGaugeName(2) = "Ferguson Falls flow"
	FlowGaugeName(3) = "Appleton flow"
	FlowGaugeName(4) = "Gordon Rapids flow"
	FlowGaugeName(5) = "Lanark stream flow"
	FlowGaugeName(6) = "Mill of Kintail flow"
	FlowGaugeName(7) = "Kinburn flow"
	FlowGaugeName(8) = "Bennett Lake outflow"
	FlowGaugeName(9) = "Dalhousie Lk outflow"
	FlowGaugeName(10) = "High Falls Flow"
	FlowGaugeName(11) = "Poole Creek at Maple Grove"
	FlowGaugeName(12) = "Carp River at Richardson"

	Const LakeOffset As Integer = FlowOffset + FlowGaugeCount + 5
	Const LakeGaugeCount As Integer = 16
	Dim LakeGaugeName(LakeGaugeCount) As String
	LakeGaugeName(0) = "Shabomeka Lake"
	LakeGaugeName(1) = "Mazinaw Lake"
	LakeGaugeName(2) = "Kashwakamak Lake"
	LakeGaugeName(3) = "Farm Lake"
	LakeGaugeName(4) = "Mississagagon Lake"
	LakeGaugeName(5) = "Big Gull Lake"
	LakeGaugeName(6) = "Crotch Lake"
	LakeGaugeName(7) = "High Falls"
	LakeGaugeName(8) = "Dalhousie Lake"
	LakeGaugeName(9) = "Palmerston Lake"
	LakeGaugeName(10) = "Canonto Lake"
	LakeGaugeName(11) = "Lanark"
	LakeGaugeName(12) = "Sharbot Lake"
	LakeGaugeName(13) = "Bennett Lake"
	LakeGaugeName(14) = "Mississippi Lake"
	LakeGaugeName(15) = "C.P. Dam"
	LakeGaugeName(16) = "Carp River at Maple Grove"

	Const WeekOffset As Integer = LakeOffset + LakeGaugeCount + 4
	Const WeekGaugeCount As Integer = 26
	Dim WeekGaugeName(WeekGaugeCount) As String
	WeekGaugeName(0) = "Shabomeka Lake (weekly)"
	WeekGaugeName(1) = "Mazinaw Lake (weekly)"
	WeekGaugeName(2) = "Little Marble Lake (weekly)"
	WeekGaugeName(3) = "Mississagagon Lake (weekly)"
	WeekGaugeName(4) = "Kashwakamak Lake (weekly)"
	WeekGaugeName(5) = "Farm Lake (weekly)"
	WeekGaugeName(6) = "Ardoch Bridge (weekly)"
	WeekGaugeName(7) = "Malcolm Lake (weekly)"
	WeekGaugeName(8) = "Pine Lake (weekly)"
	WeekGaugeName(9) = "Big Gull Lake (weekly)"
	WeekGaugeName(10) = "Buckshot Lake (weekly)"
	WeekGaugeName(11) = "Crotch Lake (weekly)"
	WeekGaugeName(12) = "High Falls G.S. (weekly)"
	WeekGaugeName(13) = "Mosque Lake (weekly)"
	WeekGaugeName(14) = "Summit Lake (weekly)"
	WeekGaugeName(15) = "Palmerston Lake (weekly)"
	WeekGaugeName(16) = "Canonto Lake (weekly)"
	WeekGaugeName(17) = "Bennett Lake (weekly)"
	WeekGaugeName(18) = "Dalhousie Lake (weekly)"
	WeekGaugeName(19) = "Silver Lake (weekly)"
	WeekGaugeName(20) = "Sharbot Lake (weekly)"
	WeekGaugeName(21) = "Widow Lake (weekly)"
	WeekGaugeName(22) = "Lanark Bridge (weekly)"
	WeekGaugeName(23) = "Lanark Dam (weekly)"
	WeekGaugeName(24) = "Almonte Bridge (weekly)"
	WeekGaugeName(25) = "Clayton Lake (weekly)"
	WeekGaugeName(26) = "C.P. Dam (weekly)"


	Dim LevelsConn As ADODB.Connection
	Set LevelsConn = New ADODB.Connection
	LevelsConn.ConnectionString = "DRIVER={MySQL ODBC 5.3 Unicode Driver};SERVER=mvc.on.ca;DATABASE=mvconc55_mvclevels;UID=mvconc55_levels1;PWD=4z9!yA;OPTION=3"
	LevelsConn.Open
	'-----------------------------------------------------------------------------------------------------------------------------'

	'The With statement ensures the macro references the daily planning cycle workbook.
	With ThisWorkbook
		For i = 0 To UBound(FlowGaugeName)
			'If the value exists, the Date, Time, Value and Historical average are uploaded.
			If .Sheets("Raw2").Range("E" & (FlowOffset + i)) < .Sheets(InputDate).Range("E" & (FlowOffset + i)) Then _
				Call Run_SQL(i + FlowOffset, InputDate, FlowGaugeName(i), LevelsConn)
		Next i

		For i = 0 To UBound(LakeGaugeName)
			If .Sheets("Raw2").Range("E" & (LakeOffset + i)) < .Sheets(InputDate).Range("E" & (LakeOffset + i)) Then _
				Call Run_SQL(i + LakeOffset, InputDate, LakeGaugeName(i), LevelsConn)
		Next i

		For i = 0 To UBound(WeekGaugeName)
			'These values are the same between the Raw2 sheet and the other sheets, so this if statement instead checks if the value is positive
			If 0 < .Sheets(InputDate).Range("E" & (WeekOffset + i)) Then _
				Call Run_SQL(i + WeekOffset, InputDate, WeekGaugeName(i), LevelsConn) 'NoRain is set to true because there is no precipitation measurement for the weekly values
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

	With ThisWorkbook
		havg = "NULL" 'Set the historical average to NULL by default

		'If there is data for the historical average, update the value of havg
		If Not IsEmpty(.Sheets(InputDate).Range("I" & i)) Then _
			havg = esc(.Sheets(InputDate).Range("I" & i))

		'The first half of the sql string is always the same; however, the second half depends on whether or not there is a precipitation value to be recorded
		strSQL = "INSERT INTO mvconc55_mvclevels.data (id, gauge, date, time, datainfo, historicalaverage, precipitation) " & _
			"VALUES (NULL, '" & esc(GaugeName) & "', '" & esc(Format(.Sheets(InputDate).Range("B" & i), "yyyy-mm-dd")) & "', '" & _
			esc(Format(.Sheets(InputDate).Range("C" & i), "h:mm AMPM")) & "', " & _
			esc(.Sheets(InputDate).Range("E" & i)) & ", "

		If IsEmpty(ThisWorkbook.Sheets(InputDate).Range("K" & i)) Or GaugeName = "Mississippi Lake" _
			  Or GaugeName = "Kashwakamak Lake" Or 0 < InStr(GaugeName, "(weekly)") Then
			'If there is a missing precipitation value (or precipitation is not recorded), then there is no precipitation value to store
			strSQL = strSQL & havg & ", NULL) ON DUPLICATE KEY UPDATE time='" & _
				esc(Format(.Sheets(InputDate).Range("C" & i), "h:mm AMPM")) & "', datainfo=" & _
				esc(.Sheets(InputDate).Range("E" & i)) & ", historicalaverage=" & havg & ", precipitation=NULL;"
		Else
			'Otherwise, there is a precipitation value and we need to add it to the sql string
			strSQL = strSQL & havg & ", " & _
				esc(.Sheets(InputDate).Range("K" & i)) & ") ON DUPLICATE KEY UPDATE time='" & _
				esc(Format(.Sheets(InputDate).Range("C" & i), "h:mm AMPM")) & "', datainfo=" & _
				esc(.Sheets(InputDate).Range("E" & i)) & ", historicalaverage=" & havg & _
				", precipitation=" & esc(.Sheets(InputDate).Range("K" & i)) & ";"
		End If

		LevelsConn.Execute strSQL
	End With
End Sub

Function esc(txt As String)
	esc = Trim(Replace(txt, "'", "\'"))
End Function