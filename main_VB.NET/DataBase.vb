Option Strict Off
Option Explicit On
Module DataBase
	
	Public cn As New ADODB.Connection
	Public rs As New ADODB.Recordset
	Public sqlstring As String
	
	
	Public Function Executesql(ByRef sqlstr As String) As Object
		Dim strPath As String
		
		On Error GoTo ADOERROR
		strPath = My.Application.Info.DirectoryPath & "\" & gstrCurProjName & "\"
		
		If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
		
		cn = New ADODB.Connection
		rs = New ADODB.Recordset
		
		rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		cn.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & strPath & "Data.mdb"
		cn.Open()
		rs.Open(sqlstr, cn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		
		Exit Function
		
ADOERROR: 
		MsgBox(Err.Source & "------" & Err.Description)
		
	End Function
End Module