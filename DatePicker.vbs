'Option Explicit

'The DPC file uses the Microsoft Date and Time Picker Control 6.0  to receive a date input from the user.
'Versions of Microsoft Excel after 2010 do not include this control in the standard installation and it therefore must be installed before the macros will work.
'To install the widget:
'1.  Check whether you are running 32 or 64 Bit Windows.  Open the file explorer and right click on 'My Computer', 'Computer' or 'This PC' depending on your version of Windows.
'    Select Properties.  The number of bits will be published under System->System Type.
'2.  Download and save the MSCOMCT2.CAB file from: https://support.microsoft.com/en-us/kb/297381 or request a copy of the file from cmcguire@mvc.on.ca.
'3.  Double click on the downloaded file.
'4.  Right click on the MSCOMCT2.OCX file and click 'Extract'.
'   a.  If using 64 bit windows extract the file to C:\Windows\SysWOW64
'   b.  If using 32 bit Windows extract the file to C:\Windows\System32
'5.  Check to make sure the file was extracted properly.
'    If the file does not extract, you will need to change the SysWOW64 or System32 folder properties.
'   a.  Right click on the SysWOW64 or System32 folder, select Properties.
'   b.  Click the Security tab
'   c.  Click the "Advanced" button.
'   d.  Click "Change" next to Owner.
'   e.  Type your username, click the "Check Names" button, then click OK.
'   f.  Check "Replace owner on subcontainers and objects" under the owner's name.
'   g.  Click OK again. If you get a message saying "Do you want to replace the directory permissions with permissions granting you full control?", click "Yes".
'   h.  Click the "Edit" button.
'   i.  Click on your username from the list.
'   j.  Check "Full control" underneath it.
'   k.  Click OK.
'   l.  Click OK again.
'6.  The MSCOMCT2.OCX file now needs to be registered.  Type cmd into the search box beside the start menu.
'    Right click on the 'Command Prompt / cmd.exe' and select 'Run as administrator'.
'7.  Check that the directory is pointed to the same folder that the MSCOMCT2.OCX file was extracted to.
'    If the folder does not match, copy "cd C:\Windows\SysWOW64" or "cd C:\Windows\System32" and press 'Enter' to update the file directory.
'8.  After the Command Prompt shows the right directory, copy and paste "regsvr32 mscomct2.ocx" into the Command Prompt and press Enter.
'9.  A message box will pop up to say if the procedure has been completed successfully.
'10. You are now ready to run the 'DatePicker' Form.


Private Sub CmdButton_Submit_Click()
	'-----------------------------------------------------------------------------------------------------------------------------'
	'Please send any questions or feedback to cmcguire@mvc.on.ca
	'-----------------------------------------------------------------------------------------------------------------------------'
	'The date submitted through the Date Picker control is used to assign values to the public variable InputDate that is defined in the DPCupdate module.
	'The format function is used to reformat the submitted date for readability in Excel.
	SheetName = Format(DatePicker.DTPicker1.Value, "mmm d")
	SheetDay = DatePicker.DTPicker1.Value + TimeSerial(6, 0, 0)

	'The form is unloaded to free up memory.
	Unload Me

End Sub

'This Private Subroutine sets the Date Picker to the current date to improve the user experience.
Private Sub UserForm_Initialize()
	DatePicker.DTPicker1.Value = Date
End Sub


