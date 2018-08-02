Private pTables As Collection

Public Sub Class_Initialize()
	Set pTables = New Collection
End Sub

Public Function Add(obj As CTable) As CTableList
	Set Add = Me

	pTables.Add obj, CStr(n)
End Function

Public Property Get count() As Integer
	count = n
End Property
Private Property Get n() As Integer
	n = pTables.count
End Property

Public Function Table(Optional i As Integer = -1) As CTable
	Set Table = Nothing
	If n - 1 < 0 Then _
		Exit Function
	If i < 0 Then
		Set Table = Table(n - 1)
		Exit Function
	End If
	Set Table = pTables(CStr(i))
End Function
