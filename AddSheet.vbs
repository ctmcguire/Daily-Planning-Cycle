Option Explicit

Sub CreateSheet(SheetName As String, InputDate As Date)
	Const NoDateCount As Integer = 2 'These 2 variables should be moved into CASpecific and renamed
	Dim unsetDate(NoDateCount-1) As String
	unsetDate(0) = "C16"
	unsetDate(1) = "C30"

	Call DebugLogging.PrintMsg("Sheet Not Found.  Adding new sheet...")

	With ThisWorkbook
		.Sheets.Add(Before:=Worksheets("Raw1")).name = SheetName

		Call DebugLogging.PrintMsg("Sheet Added.  Copying cells from Raw2")

		'The Raw2 worksheet is the template for new sheets.
		'The historicals, weekly level observations and formatting changes made on Raw2 will be pasted into all subsequent new sheets.
		.Sheets("Raw2").Range("A1:P200").Copy 

		'The Raw2 template is pasted into the new sheet. 
		.Sheets(SheetName).Range("A1").PasteSpecial xlPasteColumnWidths
		.Sheets("Raw2").Range("A3:P200").Copy Destination:=Sheets(SheetName).Range("A3")
		.Sheets("Raw2").Range("F1").Copy Destination:=Sheets(SheetName).Range("F1")
		.Sheets(SheetName).Range("F1").Formula = ThisWorkbook.Sheets("Raw2").Range("F1").Formula
		.Sheets(SheetName).Range("A3").Value = ThisWorkbook.Sheets("Raw2").Range("A3:P200").Value
		.Sheets(SheetName).Range("A3").Formula = ThisWorkbook.Sheets("Raw2").Range("A3:P200").Formula

		Call DebugLogging.PrintMsg("Cells copied.  Adding WebUpdate button (you may want to remove this next part if you don't use the WebUpdate Macro)...")

		'-----------------------------------------------------------------------------------------------------------------------------'
		'The dashboard buttons are inserted into the new sheet and formatted.
		'The button labels 'Upload2Web' and 'PrintDPC' are defined on the Raw2 sheet and can be edited to the left of the Forumla Bar.
		Dim btn As Button

		Set btn = .Sheets(SheetName).Buttons.Add(100, 5, 90, 25)
		With btn
			.OnAction = "'DPCupdate.UpdateWebBySql""" & InputDate & """'"
			.Caption = "Upload to Website"
			.name = "Upload2Web"
			.Placement = Excel.XlPlacement.xlFreeFloating
		End With

		Call DebugLogging.PrintMsg("WebUpdate button added (you may want to remove the previous part if you don't use the WebUpdate Macro)")
		Call DebugLogging.PrintMsg("Adding PrintDPC button...")

		Set btn = .Sheets(SheetName).Buttons.Add(200, 5, 90, 25)
		With btn
			.OnAction = "PrintDPC.PrintDPCPage"
			.Caption = "Print Sheet"
			.name = "PrintDPC"
			.Placement = Excel.XlPlacement.xlFreeFloating
		End With

		Call DebugLogging.PrintMsg("PrintDPC button added.  Updating TimeStamp Values...")

		'-----------------------------------------------------------------------------------------------------------------------------'
		'The date is loaded into cell B6 on the new sheet and cell formulas in the sheet populate the remaining dates.
		Range("B" & flowStart).Value = InputDate
		Range("C" & flowStart & ":C" & flowStart+flowCount & ", C" & dailyStart & ":C" & dailyStart+dailyCount).Value = TimeValue(InputDate)

		dim cell
		for each cell in unsetDate
			.Sheets(SheetName).Range(cell).Value = .Sheets("Raw2").Range(cell).Value
		next cell

		Call DebugLogging.PrintMsg("Finished updating timestamp values")
	End With
End Sub