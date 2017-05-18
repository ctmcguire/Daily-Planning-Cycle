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

'The 'levelsdata' variable references the workbook that gets uploaded to mvc.on.ca.
Dim levelsdata As Workbook
'The 'UploadSheet' variable references the worksheet that recieves values from the DPC.
Dim UploadSheet As Worksheet
'The 'i' variable is used to navigate the rows of the levelsdata workbook.
Dim i As Integer
'The 'z' variable is used to navigate the rows of the daily planning cycle workbook.
Dim z As Integer
'The 'InputDate' variable references the sheet where the Upload to Website was clicked.
Dim InputDate As String
InputDate = Format(sheetdate, "mmm d")

'-----------------------------------------------------------------------------------------------------------------------------'
Dim GaugeName(69) As String
GaugeName(6) = "Myers Cave flow"
GaugeName(7) = "Buckshot Creek flow"
GaugeName(8) = "Ferguson Falls flow"
GaugeName(9) = "Appleton flow"
GaugeName(10) = "Gordon Rapids flow"
GaugeName(11) = "Lanark stream flow"
GaugeName(12) = "Mill of Kintail flow"
GaugeName(13) = "Kinburn flow"
GaugeName(14) = "Bennett Lake outflow"
GaugeName(15) = "Dalhousie Lk outflow"
GaugeName(16) = "High Falls Flow"

GaugeName(21) = "Shabomeka Lake"
GaugeName(22) = "Mazinaw Lake"
GaugeName(23) = "Kashwakamak Lake"
GaugeName(24) = "Farm Lake"
GaugeName(25) = "Mississagagon Lake"
GaugeName(26) = "Big Gull Lake"
GaugeName(27) = "Crotch Lake"
GaugeName(28) = "Palmerston Lake"
GaugeName(29) = "Canonto Lake"
GaugeName(30) = "Sharbot Lake"
GaugeName(31) = "Bennett Lake"
GaugeName(32) = "Dalhousie Lake"
GaugeName(33) = "Lanark"
GaugeName(34) = "Mississippi Lake"
GaugeName(35) = "C.P. Dam"
GaugeName(36) = "Poole Creek at Maple Grove"
GaugeName(37) = "Carp River at Maple Grove"
GaugeName(38) = "Carp River at Richardson"
GaugeName(39) = "High Falls"

GaugeName(43) = "Shabomeka Lake (weekly)"
GaugeName(44) = "Mazinaw Lake (weekly)"
GaugeName(45) = "Little Marble Lake (weekly)"
GaugeName(46) = "Mississagagon Lake (weekly)"
GaugeName(47) = "Kashwakamak Lake (weekly)"
GaugeName(48) = "Farm Lake (weekly)"
GaugeName(49) = "Ardoch Bridge (weekly)"
GaugeName(50) = "Malcolm Lake (weekly)"
GaugeName(51) = "Pine Lake (weekly)"
GaugeName(52) = "Big Gull Lake (weekly)"
GaugeName(53) = "Buckshot Lake (weekly)"
GaugeName(54) = "Crotch Lake (weekly)"
GaugeName(55) = "High Falls G.S. (weekly)"
GaugeName(56) = "Mosque Lake (weekly)"
GaugeName(58) = "Palmerston Lake (weekly)"
GaugeName(59) = "Canonto Lake (weekly)"
GaugeName(60) = "Bennett Lake (weekly)"
GaugeName(61) = "Dalhousie Lake (weekly)"
GaugeName(62) = "Silver Lake (weekly)"
GaugeName(63) = "Sharbot Lake (weekly)"
GaugeName(64) = "Widow Lake (weekly)"
GaugeName(65) = "Lanark Bridge (weekly)"
GaugeName(66) = "Lanark Dam (weekly)"
GaugeName(67) = "Almonte Bridge (weekly)"
GaugeName(68) = "Clayton Lake (weekly)"
GaugeName(69) = "C.P. Dam (weekly)"

Dim rowCursor As Integer
Dim strSQL As String
Dim testSQL As String
Dim havg As String
Dim insavg As String
Dim LevelsConn As ADODB.Connection
Set LevelsConn = New ADODB.Connection
LevelsConn.ConnectionString = "DRIVER={MySQL ODBC 5.3 Unicode Driver};SERVER=mvc.on.ca;DATABASE=mvconc55_mvclevels;UID=mvconc55_levels1;PWD=4z9!yA;OPTION=3"
LevelsConn.Open
'-----------------------------------------------------------------------------------------------------------------------------'

'The With statement ensures the macro references the daily planning cycle workbook.
With ThisWorkbook

'This 'For' statement navigates the rows of the 'levelsdata' workbook.
For i = 6 To 69

	'This if statement determines if a flow or level value exists.
	If (i <= 29 And Not IsEmpty(.Sheets(InputDate).Range("E" & i))) Or (i = 30 And .Sheets(InputDate).Range("E30") > 190.66) Or (i = 31 And .Sheets(InputDate).Range("E31") > 151.99) Or (i = 32 And .Sheets(InputDate).Range("E32") > 150) Or (i > 32 And .Sheets(InputDate).Range("E" & i) > 0) Then
		'If the value exists, the Date, Time, Value and Historical average are uploaded.

		If IsEmpty(.Sheets(InputDate).Range("I" & i)) = True Then
		insavg = "NULL"
		havg = ", historicalaverage=NULL"
		Else
		insavg = esc(.Sheets(InputDate).Range("I" & i))
		havg = ", historicalaverage=" & esc(.Sheets(InputDate).Range("I" & i))
		End If
		

		If i <= 15 Or i = 21 Or i = 27 Or i = 28 Or (i > 29 And i < 33) Or i = 37 Then
		   
			'If statement to check if the precipitation is null. If isnull then set strSQL as the first SQL call otherwise set strSQL as the second SQL call
			If (IsEmpty(ThisWorkbook.Sheets(InputDate).Range("K" & i))) Then
				
				strSQL = "INSERT INTO mvconc55_mvclevels.data (id, gauge, date, time, datainfo, historicalaverage, precipitation) " & _
				"VALUES (NULL, '" & esc(GaugeName(i)) & "', '" & esc(Format(.Sheets(InputDate).Range("B" & i), "yyyy-mm-dd")) & "', '" & _
				esc(Format(.Sheets(InputDate).Range("C" & i), "h:mm AMPM")) & "', " & _
				esc(.Sheets(InputDate).Range("E" & i)) & ", " & _
				insavg & ", NULL) ON DUPLICATE KEY UPDATE time='" & _
				esc(Format(.Sheets(InputDate).Range("C" & i), "h:mm AMPM")) & "', datainfo=" & _
				esc(.Sheets(InputDate).Range("E" & i)) & havg & ", precipitation=NULL;"
				
				LevelsConn.Execute strSQL
				
			Else
				
				strSQL = "INSERT INTO mvconc55_mvclevels.data (id, gauge, date, time, datainfo, historicalaverage, precipitation) " & _
				"VALUES (NULL, '" & esc(GaugeName(i)) & "', '" & esc(Format(.Sheets(InputDate).Range("B" & i), "yyyy-mm-dd")) & "', '" & _
				esc(Format(.Sheets(InputDate).Range("C" & i), "h:mm AMPM")) & "', " & _
				esc(.Sheets(InputDate).Range("E" & i)) & ", " & _
				insavg & ", " & _
				esc(.Sheets(InputDate).Range("K" & i)) & ") ON DUPLICATE KEY UPDATE time='" & _
				esc(Format(.Sheets(InputDate).Range("C" & i), "h:mm AMPM")) & "', datainfo=" & _
				esc(.Sheets(InputDate).Range("E" & i)) & havg & _
				", precipitation=" & esc(.Sheets(InputDate).Range("K" & i)) & ";"
				
				LevelsConn.Execute strSQL
				
			End If
		
		Else
		
				 strSQL = "INSERT INTO mvconc55_mvclevels.data (id, gauge, date, time, datainfo, historicalaverage, precipitation) " & _
				"VALUES (NULL, '" & esc(GaugeName(i)) & "', '" & esc(Format(.Sheets(InputDate).Range("B" & i), "yyyy-mm-dd")) & "', '" & _
				esc(Format(.Sheets(InputDate).Range("C" & i), "h:mm AMPM")) & "', " & _
				esc(.Sheets(InputDate).Range("E" & i)) & ", " & _
				insavg & ", NULL) ON DUPLICATE KEY UPDATE time='" & _
				esc(Format(.Sheets(InputDate).Range("C" & i), "h:mm AMPM")) & "', datainfo=" & _
				esc(.Sheets(InputDate).Range("E" & i)) & havg & ", precipitation=NULL;"
				LevelsConn.Execute strSQL
			
		End If

	End If
	'These inline 'If' statements ensure the right row is being extracted from the daily planning cycle.
	If i = 16 Then i = i + 4
	If i = 39 Then i = i + 3
	
Next i


LevelsConn.Close


MsgBox "The requested data has been uploaded to the website. Please visit: http://mvc.on.ca/water-levels-app/levels-table-option/ataglance.php to ensure accuracy."

'The previously adjusted modes are returned to their default state.
Application.StatusBar = False
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End With
End Sub


Function esc(txt As String)
	esc = Trim(Replace(txt, "'", "\'"))
End Function

