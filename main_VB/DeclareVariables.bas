Attribute VB_Name = "Module3"
Option Explicit

Public Const FixR = 1
Public Const FixG = 2
Public Const FixB = 3

Public Const AdjustSingle = 1
Public Const AdjustDouble = 0

Public Const SingleStep = 0
Public Const ComplexStep = 1

Public Const HighBri = 1
Public Const LowBri = 0

Type ColorTemp
    x As Single
    y As Single
    lv As Single
End Type

Public strBuff As String

Public i As Integer


Public Const xxf = 1
Public Const xfyf = 2
Public Const yyf = 3
Public Const microStep = True
Public Const StepbyStep = False



Public IsCa210Channel As Long
Public IsStepTime As Long
Public IsWhitePtn As String

Public IsAdjCool_1 As Boolean
Public IsAdjCool_2  As Boolean
Public IsAdjNormal  As Boolean
Public IsAdjWarm_1 As Boolean
Public IsAdjWarm_2  As Boolean

Public IsAdj5400k As Boolean
Public IsAdj5000k As Boolean
Public IsAdj4000k As Boolean
Public IsAdj2600k As Boolean

Public IsBarcodeLen As Integer
Public IsFunctionAutoBri As Boolean
Public IsSensorLight As Boolean
Public IsSaveData As Boolean
Public IsCheckColorTemp  As Boolean

Public IsSendOffset As Boolean
Public IsAdjsutOffset As Boolean

Public strCurrentModelName As String
Public strDataVersion As String
Public IsStop As Boolean
Public IsACK As Boolean
Public SetTVCurrentComID As Integer
Public SetData As Integer
Public SetDay As Integer
Public IsCool_1ModeIndex As Long
Public IsCool_2ModeIndex As Long
Public IsNormalModeIndex As Long
Public IsWarm_1ModeIndex As Long
Public IsWarm_2ModeIndex As Long

Public IsCa210ok As Boolean

Public IsSNWriteSuccess As Boolean
Public scanbarcode As String
Public strSerialNo As String
Public countTime As Long
Public DebugFlag As Boolean

Public SetTVCurrentComBaud As Long


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

Public Const valColorTempCool1 As Long = 12000
Public Const valColorTempNormal As Long = 10000
Public Const valColorTempWarm1 As Long = 6500
