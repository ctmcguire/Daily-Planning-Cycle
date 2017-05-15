Option Explicit

'The date picker assigns a value to the public variable 'InputDay'.
'The variable is public so that it can be used by all the called functions.
Public InputDay As String
Public InputNumber As Date
'-------------------------------------------------------------------------------------------------------------------------------------------------'
'The UpdateDPC subroutine is run when the Update DPC button is pressed on sheet 'Raw1'
'UpdateDPC runs the modules that load a new sheet and populate it with the requested data.
Sub UpdateDPC()
'-------------------------------------------------------------------------------------------------------------------------------------------------'
'The 'DPCsheet' variable is used to check if the requested sheet already exists.
Dim DPCsheet As Excel.Worksheet

'The Status Bar is located on the bottom left corner of the Excel window.  It's default status is 'READY'.
'The Status Bar Displays 'Processing Request...' until the UpdateDPC subroutine has ended.
Application.StatusBar = "Processing Request..."
'Screen Updating is turned off to speed up the processing time.
Application.ScreenUpdating = False
'Sheet calculations are turned off to speed up the processing time.
Application.Calculation = xlCalculationManual

'The Date Picker form is shown and the user inputs a date.
'Instructions to install the Date Picker control can be found by right clicking on the 'DatePicker' form and selecting 'View Code'.
DatePicker.Show

Dim Answer As Integer

'This 'For' loop checks that the a sheet for the requested date does not already exist.
For Each DPCsheet In ThisWorkbook.Worksheets
    If DPCsheet.name = InputDay Then
        'If a sheet with the requested date already exists, the subroutine exits so that previous data is not overwritten.
        Answer = MsgBox("A sheet for '" & InputDay & "' already exists.  Do you want to fill empty cells?", vbYesNo, "Sheet Already Exists")
        If Answer = vbYes Then
            DPCsheet.Unprotect
        Else
            Exit Sub
        End If
    End If
Next


'----------------------------------------------------------------------------------------------------------------------------------------------'
'The 'AddSheet' module creates a new sheet in the workbook, names it after the requested date, and pastes the template from 'Raw2'.
If Answer = 0 Then Call AddSheet.CreateSheet(InputDay, InputNumber)
'The KiWISLoader module loads the KiWIS tables to the sheet 'Raw1'.
Call KiWISLoader.KiWIS_Import(InputNumber)
'The KiWIS2Excel module loads the data from Raw1 to the new sheet that was created.
Call KiWIS2Excel.Raw1Import(InputDay)

'The Weather... modules scrape weather data from AccuWeather, Environment Canada and The Weather Network and pastes it into the new sheet.
If Answer = 0 Then
Call WeatherAccu.AccuWeatherScraper(InputDay)
Call WeatherEC.ECWeatherScraper(InputDay)
Call WeatherTWN.TWNWeatherScraper(InputDay)
End If

'The DataProtector module locks cells for editing and saves a backup of the daily planning cycle to the local desktop and the Water Management Files folder.
Call DataProtector.LockCells(InputDay, InputNumber)

'----------------------------------------------------------------------------------------------------------------------------------------------'
'This 'With' statement ensures that the workbook is open on the new sheet after the module has run.
With ThisWorkbook.Worksheets(InputDay)
    'The Activate function is not recommended.  This line may cause execution errors.
    .Activate
    .Range("E16").Select
End With

MsgBox "The requested sheet for " & InputDay & " has loaded."

'The previously adjusted modes are returned to their default state.
Application.StatusBar = False
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub
