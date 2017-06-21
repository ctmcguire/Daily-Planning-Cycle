Sub DailyEmail()
	Dim WB As Workbook
	Dim FileName As String
	Dim date1 As String
	Dim imsg As Object
	Dim iconf As Object
	Dim flds As Object
	Dim schema As String

	'turn of unnecessary applications to speed up process
	Application.Calculation = xlManual
	With Application
		.EnableEvents = False
		.ScreenUpdating = False
	End With

	'date1 is set as the current date
	date1 = Format(Date, "mmm d")
	Set WB = Application.ActiveWorkbook
	FileName = WB.FullName

	'change the page setup, this way the pdf is formatted clearly
	Application.PrintCommunication = False
	With ActiveSheet.PageSetup
		.Orientation = xlPortrait
		.Zoom = False
		.PrintArea = Worksheets(date1).Range("A1:P123")
		.LeftMargin = Application.InchesToPoints(0.25)
		.RightMargin = Application.InchesToPoints(0.25)
		.TopMargin = Application.InchesToPoints(0.25)
		.BottomMargin = Application.InchesToPoints(0.25)
		.HeaderMargin = Application.InchesToPoints(0)
		.FooterMargin = Application.InchesToPoints(0)
		.FitToPagesWide = 1
		.FitToPagesTall = 1
	End With
	Application.PrintCommunication = True

	'set the file name of the pdf
	xIndex = VBA.InStrRev(FileName, ".")
	If 1 < xIndex Then _
		FileName = VBA.Left(FileName, xIndex - 1) 'name of the pdf will be the title of the workbook along with the name of the current sheet
	FileName = FileName & "_" + Worksheets(date1).name & ".pdf"

	Worksheets(date1).ExportAsFixedFormat Type:=xlTypePDF, FileName:=FileName'export the current sheet as a pdf
	
	Set imsg = CreateObject("CDO.Message")
	Set iconf = CreateObject("CDO.Configuration")
	Set flds = iconf.Fields

	' send one copy with SMTP server (with autentication)
	schema = "http://schemas.microsoft.com/cdo/configuration/"
	flds.Item(schema & "sendusing") = 2 'Using port
	flds.Item(schema & "smtpserver") = "mail.mvc.on.ca"
	flds.Item(schema & "smtpserverport") = 25
	flds.Item(schema & "smtpauthenticate") = cdoBasic
	flds.Item(schema & "sendusername") = "water-management@mvc.on.ca"
	flds.Item(schema & "sendpassword") = "waterwater"
	flds.Item(schema & "smtpusessl") = False
	flds.Update

	'details of the email sent to water-management@mvc.on.ca or cmcguire@mvc.on.ca or whyyouchris@gmail.com
	With imsg
		.To = "cmcguire@mvc.on.ca"
		.From = "water-management@mvc.on.ca"
		.Sender = "DPC System"
		.Subject = "Daily Update for " + date1
		.HTMLBody = "" 'Need body or attachments will get corrupted
		'attaches the pdf file
		.AddAttachment FileName
		
		'uncomment the follwoing to attach the inserted jpg to email:
		'.Attachments.Add "C:\Users\water-management\Desktop\Jpgs\Daily.jpg"

		
		'either display the email before being sent or send immediately
		Set .Configuration = iconf
'		.send
		
	End With
	'delete the pdf
	Kill FileName

	With Application
			.EnableEvents = True
			.ScreenUpdating = True
		End With
	Application.Calculation = xlCalculationAutomatic

End Sub