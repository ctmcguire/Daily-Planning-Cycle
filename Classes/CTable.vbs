Private pStart As Integer
Private pLength As Integer
Private pGauges As Collection

Public Sub Class_Initialize()
	Set pGauges = new Collection
End Sub

Public Function Init(Row As Integer, Optional Length As Integer = -1) As CTable
	Dim temp As New CTable
	Set Init = temp.CTable(Row, Length)
End Function
Public Function CTable(Row As Integer, Optional Length As Integer = -1) As CTable
	Set CTable = Me

	pStart = Row
	pLength = Length
End Function

Public Property Get row() As Integer
	row = pStart
End Property

Public Property Get count() As Integer
	count = n
End Property
Private Property Get n() As Integer
	n = pGauges.Count
End Property

Public Function Add(ParamArray Gauges() As Variant) As CTable
	Set Add = Me

	Dim Gauge As Variant
	For Each Gauge In Gauges
		If 0 < pLength And pLength <= n Then _
			Exit Function
		pGauges.Add Gauge, CStr(n)
	Next
End Function

Public Function LoadData(SheetName As String, Optional IsAuto As Boolean = False) As CTable
	Set LoadData = Me

	Dim i As Integer
	i = 0
	Dim Gauge
	For Each Gauge In pGauges
		Gauge.LoadData SheetName, pStart + i, IsAuto
		i = i + 1
	Next
End Function

Public Function RunSql(InputDate As String, LevelsConn As ADODB.Connection) As CTable
	Set RunSql = Me

	Dim i As Integer
	For i = 0 To n - 1
		If ThisWorkbook.Sheets("Raw2").Range("E" & (pStart + i)) < ThisWorkbook.Sheets(InputDate).Range("E" & (pStart + i)) Then _
			pGauges(CStr(i)).RunSql (pStart + i), InputDate, LevelsConn
	Next
End Function