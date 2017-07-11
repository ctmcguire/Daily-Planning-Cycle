'CGaugeSensor Class
Private pName As String 'What value the gauge sensor measures (flow, level, precipitation, etc)
Private pColumn As String 'Column where this sensor's data will appear in the table
Private pRangeIndex As Integer 'Column where this sensor's data is retrieved from in raw1

Private pInitialized


Public Sub Class_Initialize()
	pInitialized = False
End Sub

' * The CGaugeSensor Function is used to initialize the values in a new CGaugeSensor Object in place of its constructor.
' * This is due mostly to the fact that VBA does not support constructors with parameters, resulting in the 
' * need for this function.
' * 
' * @param Name   - The name of what this Sensor measures
' * @param Column - The letter of the column for this sensor in the dpc tables
' * @param RawCol - The letter of the column for this sensor in the raw1 table
' * 
' * @returns - This function does not return anything
' * 
' * 
' * Example usage:
' * 				'These first 2 lines are shown for context
' * 				'Dim Sensor As CGaugeSensor
' * 				'Set Sensor = New CGaugeSensor
' * 				Sensor.CGaugeSensor "Dave the intern", "D", "I"
' * The above example initializes the CGaugeSensor Sensor with a Name of "Dave the intern", a Column of "D", and a RawCol of "I"
'**/
Public Sub CGaugeSensor(Name As String, Column As String, RangeIndex As Integer)
	If pInitialized Then _
		Exit Sub
	pName = Name
	pColumn = Column
	pRangeIndex = RangeIndex'Need to strip off "'Raw1'!" and the $'s

	pInitialized = True
End Sub

Public Property Get Column()
	Column = pColumn
End Property

Public Property Get Name()
	Name = pName
End Property

Public Function Value(ID As String)
	Dim Range As String

	Range = GetRange()

	If pName = RainName Then
		Value = Sum(ID, Range)
		Exit Function
	End If
	Value = GetData(ID, Range)
End Function

Private Function GetRange()
	Dim Range As String
	Dim Colon As Integer
	Dim Column As String
	
	Range = ThisWorkbook.Sheets("Raw1").Names("ExternalData_" & pRangeIndex)
	Range = Replace(Range, "='Raw1'!", "")
	Range = Replace(Range, "$", "")
	Colon = InStr(Range, ":")
	Column = Left(Right(Range, Len(Range) - Colon), 1)
	GetRange = Column & Right(Range, Len(Range) - 1)
End Function

Private Function GetData(ID As String, Range As String)
	GetData = Application.WorksheetFunction.Index(ThisWorkbook.Sheets("Raw1").Range(Range), (Application.WorksheetFunction.Match(ID, ThisWorkbook.Sheets("Raw1").Range(Range), 0) + 5))
End Function

Private Function Sum(ID As String, Range As String)
	Dim Column As String
	Column = Left(Range, 1)
	With ThisWorkbook.Sheets("Raw1")
		If Not (.Range(Column & (Application.WorksheetFunction.Match(ID, .Range(Range), 0) + 3))) = 7 Then
			Sum = ""
			Exit Function
		End If

		Dim Row As Integer
		Row = Application.WorksheetFunction.Match(ID, .Range(Range), 0) + 6
		Sum = Application.WorksheetFunction.Sum(.Range(Column & Row, Column & Row+12))
	End With
End Function