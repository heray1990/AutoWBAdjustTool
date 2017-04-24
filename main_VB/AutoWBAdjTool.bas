Attribute VB_Name = "AutoWBAdjTool"
Option Explicit

Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public sqlstring As String


Public Sub Main()
    FormSplash.Show
End Sub

Public Function FuncOpenSQL(sqlstr As String)
    On Error GoTo ADOERROR
    Dim strPath As String
    strPath = App.Path
    
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

