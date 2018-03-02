Function DailyEmail()
	Dim WB As Workbook
	Dim FileName As String
	Dim date1 As String
	Dim imsg As Object
	Dim iconf As Object
	Dim flds As Object
	Dim schema As String

	Call DebugLogging.Clear
	On Error Goto OnError
	Call DebugLogging.PrintMessage "Getting workbook name..."

	'date1 is set as the current date
	date1 = Format(Date, "mmm d")
	If ThisWorkbook.Sheets(date1).Range("H1").Value <> "" Then
		DailyEmail = DebugLogging.PrintMessage
		Exit Function
	End If
	Set WB = Application.ActiveWorkbook
	FileName = WB.FullName

	Call DebugLogging.PrintMessage "Changing page setup..."
	'change the page setup, this way the pdf is formatted clearly
	Application.PrintCommunication = False
	With ThisWorkbook.Sheets(date1).PageSetup
		.Orientation = xlPortrait
		.Zoom = False
		.FitToPagesTall = 1
		.FitToPagesWide = 1
		.PaperSize = xlPaperLegal
		.BlackAndWhite = False
		.LeftMargin = Application.CentimetersToPoints(0)
		.RightMargin = Application.CentimetersToPoints(0)
		.TopMargin = Application.CentimetersToPoints(0)
		.BottomMargin = Application.CentimetersToPoints(0)
		.HeaderMargin = Application.CentimetersToPoints(0)
		.FooterMargin = Application.CentimetersToPoints(0)
	End With
	Application.PrintCommunication = True

	Call DebugLogging.PrintMessage "Getting pdf name..."
	'set the file name of the pdf
	xIndex = VBA.InStrRev(FileName, ".")
	If 1 < xIndex Then _
		FileName = VBA.Left(FileName, xIndex - 1) 'name of the pdf will be the title of the workbook along with the name of the current sheet
	FileName = FileName & "_" + Worksheets(date1).name & ".pdf"

	Call DebugLogging.PrintMessage "Creating pdf nile..."
	Worksheets(date1).ExportAsFixedFormat Type:=xlTypePDF, FileName:=FileName 'export the current sheet as a pdf
	
	Set imsg = CreateObject("CDO.Message")
	Set iconf = CreateObject("CDO.Configuration")
	Set flds = iconf.Fields

	Call DebugLogging.PrintMessage "Setting email configuration..."
	' send one copy with SMTP server (with autentication)
	schema = "http://schemas.microsoft.com/cdo/configuration/"
	flds.Item(schema & "sendusing") = 2 'Using port
	flds.Item(schema & "smtpserver") = StrVal(ThisWorkbook.Names("EmailSv"))
	flds.Item(schema & "smtpserverport") = 25
	flds.Item(schema & "smtpauthenticate") = cdoBasic
	flds.Item(schema & "sendusername") = StrVal(ThisWorkbook.Names("EmailUn"))
	flds.Item(schema & "sendpassword") = StrVal(ThisWorkbook.Names("EmailPw"))
	flds.Item(schema & "smtpusessl") = False
	flds.Update

	Call DebugLogging.PrintMessage "Setting email options..."
	'details of the email sent to water-management@mvc.on.ca or cmcguire@mvc.on.ca
	With imsg
		.To = Recipients
		.from = StrVal(ThisWorkbook.Names("EmailUn"))
		.Sender = "DPC System"
		.Subject = "Daily Update for " + date1
		.HTMLBody = "" 'Need body or attachments will get corrupted
		'attaches the pdf file
		.AddAttachment FileName
		Set .Configuration = iconf

		Call DebugLogging.PrintMessage "Sending email..."
		On Error Resume Next
		.send
		If Err.Number = 0 Then
			Call DebugLogging.PrintMessage "Email sent successfully!"
			Call DebugLogging.PrintMessage "Marking worksheet as sent..."
			ThisWorkbook.Sheets(date1).Range("H1").Value = "Email sent at " & Now
			ThisWorkbook.Save
		End If
		On Error GoTo OnError 'Go back to using to default error handler
	End With
	Call DebugLogging.PrintMessage "Deleting created pdf file..."
	'delete the pdf
	Kill FileName

	DailyEmail = DebugLogging.PrintMessage
	Exit Function
	OnError:
		Call DebugLogging.Erred
		DailyEmail = DebugLogging.PrintMessage
End Function


Function StrVal(Formula As String) As String
	StrVal = Mid(Formula, 3, Len(Formula) - 3)
End Function