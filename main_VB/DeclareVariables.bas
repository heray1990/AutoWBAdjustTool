Attribute VB_Name = "DeclareVariables"
Option Explicit

Public Const FixR As Integer = 1
Public Const FixG As Integer = 2
Public Const FixB As Integer = 3
Public Const adjustMode1 As Integer = 1
Public Const adjustMode2 As Integer = 2
Public Const adjustMode3 As Integer = 3
Public Const adjustMode4 As Integer = 4

Public Const AdjustSingle As Integer = 1
Public Const AdjustDouble As Integer = 0

Public Const SingleStep As Integer = 0
Public Const ComplexStep As Integer = 1

Public Const HighBri As Integer = 1
Public Const LowBri As Integer = 0

Public Const cstrColorTempCool1 As String = "COOL1"
Public Const cstrColorTempNormal As String = "NORMAL"
Public Const cstrColorTempWarm1 As String = "WARM1"

Public Const lastChkShwDataStep As Integer = 6
Public Const cmdReceiveWaitS As Integer = 5
Public Const strRemoteHost As String = "192.168.1.11"
Public Const lngRemotePort As Long = 8888

Type ColorTemp
    x As Single
    y As Single
    lv As Single
End Type

Public strBuff As String

Public i As Integer
Public adjustGainAgainCool1Flag As Integer
Public adjustGainAgainNormalFlag As Integer
Public adjustGainAgainWarm1Flag As Integer

Public Const xxf = 1
Public Const xfyf = 2
Public Const yyf = 3
Public Const microStep = True
Public Const StepbyStep = False

Public Ca210ChannelNO As Long
Public delayTime As Long

Public isAdjustCool1 As Boolean
Public isAdjustCool2  As Boolean
Public isAdjustNormal  As Boolean
Public isAdjustWarm1 As Boolean
Public isAdjustWarm2  As Boolean

Public gintBarCodeLen As Integer
Public IsFunctionAutoBri As Boolean
Public isCheckColorTemp  As Boolean

Public isAdjustOffset As Boolean

Public gstrCurProjName As String
Public IsStop As Boolean
Public IsACK As Boolean
Public setTVCurrentComID As Integer
Public setTVCurrentComBaud As Long
Public setTVInputSource As String
Public setTVInputSourcePortNum As Integer
Public maxBrightnessSpec As Long

Public IsCa210ok As Boolean
Public isNetworkConnected As Boolean
Public utdCommMode As CommunicationMode

Public gstrBarCode As String
Public countTime As Long
Public gstrBrand As String

Public gstrVPGModel As String
Public glngI2cClockRate As Long
Public gstrVPGTiming As String
Public gstrVPG100IRE As String
Public gstrVPG80IRE As String
Public gstrVPG20IRE As String


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
Public rRGB As REALRGB
Public rRGB1 As REALRGB

Public Enum CommunicationMode
    modeUART = 1
    modeNetwork
    modeI2c
End Enum

Public Sub Log_Info(strLog As String)
    Form1.CheckStep.Text = Form1.CheckStep.Text + strLog + vbCrLf
    Form1.CheckStep.SelStart = Len(Form1.CheckStep)

    SaveLogInFile strLog
End Sub

Public Sub SaveLogInFile(strLog As String)
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
