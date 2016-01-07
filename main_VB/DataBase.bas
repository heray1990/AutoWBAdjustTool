Attribute VB_Name = "Module6"
Option Explicit

Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public sqlstring As String


Public Function Executesql(sqlstr As String)
Dim strPath As String
On Error GoTo ADOERROR
strPath = App.path
If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
cn.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & strPath & "Data.mdb"
cn.Open
rs.Open sqlstr, cn, adOpenDynamic, adLockOptimistic
Exit Function

ADOERROR:
    MsgBox Err.Source & "------" & Err.Description
   
End Function

Public Function Executesql3(sqlstr As String)
Dim strPath As String
Dim strSerialNo As Long
Dim a, b, c, d As Long
a = 20
b = 5
c = 6
d = 1
On Error GoTo ADOERROR
strPath = App.path
If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
cn.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & strPath & "Data.mdb"
cn.Open
rs.Open "select * from DataRecord", cn, adOpenDynamic, adLockOptimistic
rs.AddNew

rs.Fields(0) = strCurrentModelName

rs.Fields(1) = strSerialNo
rs.Fields(2) = b
rs.Fields(3) = c
rs.Fields(4) = d
rs.Update


Exit Function

ADOERROR:
    MsgBox Err.Source & "------" & Err.Description
   
End Function



