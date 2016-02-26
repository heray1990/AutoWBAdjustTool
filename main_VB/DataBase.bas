Attribute VB_Name = "DataBase"
Option Explicit

Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public sqlstring As String


Public Function Executesql(sqlstr As String)
    Dim strPath As String

On Error GoTo ADOERROR
    strPath = App.Path & "\" & gstrCurProjName & "\"
    
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

