Attribute VB_Name = "AutoWBAdjTool"
Option Explicit

'==========User-defined Type==========
Public Type udtConfigData
    strModel As String
    CommMode As CommunicationMode
    strComBaud As String
    intComID As Integer
    strInputSource As String
    lngDelayMs As Long
    intChannelNum As Integer
    intBarCodeLen As Integer
    intLvSpec As Integer
    strVPGModel As String
    strVPGTiming As String
    strVPG100IRE As String
    strVPG80IRE As String
    strVPG20IRE As String
    lngI2cClockRate As Long
    bolEnableCool2 As Boolean
    bolEnableCool1 As Boolean
    bolEnableNormal As Boolean
    bolEnableWarm1 As Boolean
    bolEnableWarm2 As Boolean
    bolEnableChkColor As Boolean
    bolEnableAdjOffset As Boolean
    strChipSet As String
End Type

Public Type udtSpecData
    intSPECCool1x As Long
    intSPECCool1y As Long
    intSPECCool1Lv As Long
    intSPECNormalx As Long
    intSPECNormaly As Long
    intSPECNormalLv As Long
    intSPECWarm1x As Long
    intSPECWarm1y As Long
    intSPECWarm1Lv As Long
    intTOLCool1xt As Long
    intTOLCool1yt As Long
    intTOLNormalxt As Long
    intTOLNormalyt As Long
    intTOLWarm1xt As Long
    intTOLWarm1yt As Long
    intCHKCool1Cxt As Long
    intCHKCool1Cyt As Long
    intCHKNormalCxt As Long
    intCHKNormalCyt As Long
    intCHKWarm1Cxt As Long
    intCHKWarm1Cyt As Long
    intPRESETGANCool1R As Long
    intPRESETGANCool1G As Long
    intPRESETGANCool1B As Long
    intPRESETGANNormalR As Long
    intPRESETGANNormalG As Long
    intPRESETGANNormalB As Long
    intPRESETGANWarm1R As Long
    intPRESETGANWarm1G As Long
    intPRESETGANWarm1B As Long
    intPRESETOFFCool1R As Long
    intPRESETOFFCool1G As Long
    intPRESETOFFCool1B As Long
    intPRESETOFFNormalR As Long
    intPRESETOFFNormalG As Long
    intPRESETOFFNormalB As Long
    intPRESETOFFWarm1R As Long
    intPRESETOFFWarm1G As Long
    intPRESETOFFWarm1B As Long
    intCLEVELRGBGMin As Long
    intCLEVELRGBGMax As Long
    intCLEVELRGBOMin As Long
    intCLEVELRGBOMax As Long
    intMAGICVALGMin As Long
    intMAGICVALGMax As Long
    intMAGICVALOMin As Long
    intMAGICVALOMax As Long
End Type

Public Type COLORTEMPSPEC
    xx                         As Long
    yy                         As Long
    lv                         As Long
    nColorRR                   As Long
    nColorGG                   As Long
    nColorBB                   As Long
    xt                         As Long
    yt                         As Long
    nLowRR                     As Long
    nLowGG                     As Long
    nLowBB                     As Long
End Type

Public Type REALCOLOR
    xx                         As Long
    yy                         As Long
    lv                         As Long
End Type

Public Type REALRGB
    cRR                        As Long
    cGG                        As Long
    cBB                        As Long
End Type

Public Enum CommunicationMode
    modeUART = 1
    modeNetwork
    modeI2c
End Enum

'==========Const==========
Public Const ADJMODE_1 As Integer = 1
Public Const ADJMODE_2 As Integer = 2
Public Const ADJMODE_3 As Integer = 3
Public Const ADJMODE_4 As Integer = 4

Public Const ADJMODE_GAIN As Integer = 1
Public Const ADJMODE_OFFSET As Integer = 0

Public Const COLORTEMP_COOL1 As String = "COOL1"
Public Const COLORTEMP_STANDARD As String = "NORMAL"
Public Const COLORTEMP_WARM1 As String = "WARM1"

Public Const LASTSTEP As Integer = 6
Public Const REMOTE_HOST As String = "192.168.1.11"
Public Const REMOTE_PORT As Long = 8888

'==========Public Variables==========
Public gEnumCommMode As CommunicationMode
Public gudtConfigData As udtConfigData
Public gudtSpecData As udtSpecData

Public glngCaChannel As Long
Public glngDelayTime As Long
Public gintCurComBaud As Long
Public glngBlSpecVal As Long
Public glngI2cClockRate As Long

Public gintBarCodeLen As Integer
Public gintCurComId As Integer
Public gintTvInputSrcPort As Integer

Public gblnEnableCool1 As Boolean
Public gblnEnableCool2  As Boolean
Public gblnEnableStandard  As Boolean
Public gblnEnableWarm1 As Boolean
Public gblnEnableWarm2  As Boolean
Public gblnChkColorTemp  As Boolean
Public gblnAdjOffset As Boolean
Public gblnStop As Boolean
Public gblnCaConnected As Boolean
Public gblnNetConnected As Boolean

Public gstrChipSet As String
Public gstrCurProjName As String
Public gstrTvInputSrc As String
Public gstrVPGModel As String
Public gstrVPGTiming As String
Public gstrVPG100IRE As String
Public gstrVPG80IRE As String
Public gstrVPG20IRE As String

Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset


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

Public Sub SubLogInfo(strLog As String)
    FormMain.CheckStep.Text = FormMain.CheckStep.Text + strLog + vbCrLf
    FormMain.CheckStep.SelStart = Len(FormMain.CheckStep)

    SubSaveLogInFile strLog
End Sub

Public Sub SubSaveLogInFile(strLog As String)
    Dim logPath As String

    logPath = App.Path & "\" & "Logs\"
    If Right(logPath, 1) <> "\" Then logPath = logPath & "\"
    
    If Dir(logPath, vbDirectory) = "" Then
        MkDir logPath
    End If
    
    Open (logPath & gstrCurProjName & "-" & Format(Date, "YYYY-MM-DD") & ".log") For Append As #1
    Print #1, CStr(Time) & "> " & strLog
    Close #1
End Sub

Public Sub SubDelayMs(mmSec As Long)
    On Error GoTo ShowError
    Dim start As Single

    start = Timer
    While (Timer - start) < (mmSec / 1000#)
        DoEvents
   
        If gblnStop = True Then
            Exit Sub
        End If
    Wend
    Exit Sub

ShowError:
    MsgBox Err.Source & "------" & Err.Description
End Sub

Public Sub SubDelayWithFlag(Sec As Long, flag As Boolean)
    On Error GoTo ShowError
    Dim start As Single

    start = Timer
    While (Timer - start) < Sec
        DoEvents
   
        If flag = True Then
            Exit Sub
        End If
        
        If gblnStop = True Then
            Exit Sub
        End If
    Wend
    Exit Sub

ShowError:
    MsgBox Err.Source & "------" & Err.Description
End Sub
