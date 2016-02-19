VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto Color Temp Adjust System"
   ClientHeight    =   4620
   ClientLeft      =   5865
   ClientTop       =   2625
   ClientWidth     =   10335
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   10560
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      DrawWidth       =   3
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   2640
      Picture         =   "Form1.frx":1DF72
      ScaleHeight     =   2550
      ScaleWidth      =   3780
      TabIndex        =   6
      Top             =   960
      Width           =   3810
      Begin VB.Label lbColorTempWrong 
         BackStyle       =   0  'Transparent
         Caption         =   "Out Range"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   10
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10560
      Top             =   3480
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   10560
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox CheckStep 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3580
      Left            =   6440
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Form1.frx":3D06C
      Top             =   960
      Width           =   3805
   End
   Begin VB.TextBox txtInput 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "123456789"
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label lbCommMode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UART"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   120
      TabIndex        =   21
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "2970"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4480
      TabIndex        =   20
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lbModelName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sampl1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lbTimer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "0s"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   5640
      TabIndex        =   18
      Top             =   4130
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "2670"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3120
      TabIndex        =   17
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SPEC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   2640
      TabIndex        =   16
      Top             =   4020
      Width           =   3805
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "INITIAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   120
      TabIndex        =   15
      Top             =   4020
      Width           =   2535
   End
   Begin VB.Label lbAdjustWARM_2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WARM2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   120
      TabIndex        =   14
      Top             =   3510
      Width           =   2535
   End
   Begin VB.Label lbAdjustCOOL_2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COOL2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   120
      TabIndex        =   13
      Top             =   1980
      Width           =   2535
   End
   Begin VB.Label Label_Lv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "210"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   5520
      TabIndex        =   12
      Top             =   3525
      Width           =   930
   End
   Begin VB.Label Label_y 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "2800"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   525
      Left            =   4500
      TabIndex        =   11
      Top             =   3555
      Width           =   975
   End
   Begin VB.Label Label_x 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "2700"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   525
      Left            =   3120
      TabIndex        =   10
      Top             =   3555
      Width           =   975
   End
   Begin VB.Label lbAdjustWARM_1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WARM1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label lbAdjustNormal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NORMAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   120
      TabIndex        =   3
      Top             =   2490
      Width           =   2535
   End
   Begin VB.Label lbAdjustCOOL_1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COOL1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   120
      TabIndex        =   2
      Top             =   1470
      Width           =   2535
   End
   Begin VB.Label checkResult 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ADJUST COLOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   38.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2640
      TabIndex        =   0
      Top             =   0
      Width           =   7605
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   x:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   2640
      TabIndex        =   8
      Top             =   3525
      Width           =   1545
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " y:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   4170
      TabIndex        =   9
      Top             =   3525
      Width           =   1425
   End
   Begin VB.Menu vbFunc 
      Caption         =   "Function"
      Begin VB.Menu vbConCA310 
         Caption         =   "ConnectCA210"
      End
      Begin VB.Menu tbDisConnectastro 
         Caption         =   "DisConnectCA210(&D)"
      End
      Begin VB.Menu vbConChroma 
         Caption         =   "ConnectChroma"
      End
   End
   Begin VB.Menu vbSet 
      Caption         =   "Setting"
      Begin VB.Menu vbSetSPEC 
         Caption         =   "Set Spec"
      End
   End
   Begin VB.Menu vbDescription 
      Caption         =   "Description"
      Begin VB.Menu vbAbout 
         Caption         =   "About"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim RES As Long
Dim Result As Boolean
Dim presetData As COLORTEMPSPEC
Dim c12000K As COLORTEMPSPEC
Dim c10000K As COLORTEMPSPEC
Dim c6500K As COLORTEMPSPEC
Dim cFF12000K As COLORTEMPSPEC
Dim cFF10000K As COLORTEMPSPEC
Dim cFF6500K As COLORTEMPSPEC
Dim rColor As REALCOLOR
Dim rColorLastChk As REALCOLOR
Dim Calibrate, MinBrightness As Long
Dim resCodeForAdjustColorTemp As Long
Dim cmdMark As String
Dim clsCommProtocal As CommunicationProtocal
Dim clsHaierProtocal As HaierProtocal
Dim clsLetvProtocal As LetvProtocal

Private Sub subMainProcesser()
    Dim i, j As Integer

On Error GoTo ErrExit
    subInitBeforeRunning

    If IsStop = True Then
        Exit Sub
    End If
    
    If IsSNWriteSuccess = funSNWrite Then
        If IsStop = True Then
            Exit Sub
        End If
        
        txtInput = scanbarcode
    Else
        ShowError_Sys (6)
        GoTo FAIL
    End If
    
On Error GoTo ErrExit

    If IsCa210ok = False Then
        MsgBox "CA210 disconnected,Please click'Connect'->'Connect CA210'to do operation!", vbOKOnly + vbInformation, "warning"
        subInitAfterRunning
        
        Exit Sub
    End If

    checkResult.BackColor = &H80FFFF
    IsStop = False
    checkResult.Caption = "RUN..."
    checkResult.ForeColor = &HC0&
    CheckStep = ""
    CheckStep.BackColor = &H8000000F
    CheckStep.ForeColor = &H80000008

    lbAdjustCOOL_1.BackColor = &H8000000F
    lbAdjustCOOL_2.BackColor = &H8000000F
    lbAdjustNormal.BackColor = &H8000000F
    lbAdjustWARM_1.BackColor = &H8000000F
    lbAdjustWARM_2.BackColor = &H8000000F

    Picture1.Cls
    lbColorTempWrong.Visible = False

    Set ObjMemory = ObjCa.Memory
    ObjMemory.ChannelNO = Ca210ChannelNO

    strBuff = ""

    Log_Info "###INITIAL USER###"
    Log_Info "###ADJUST COLORTEMP###"

    'ENTER_FAC_MODE
    clsCommProtocal.EnterFacMode
    DelayMS delayTime
    
    'Call SEL_INPUT_SOURCE_FOR_WB(setTVInputSource, setTVInputSourcePortNum)
    Call clsCommProtocal.SwitchInputSource(setTVInputSource, setTVInputSourcePortNum)
    DelayMS delayTime
    
    If Not (gstrBrand = "Letv") Then
        Call clsCommProtocal.SetBrightness(50)
        DelayMS delayTime
        
        Call clsCommProtocal.SetContrast(50)
        DelayMS delayTime
        
        Call clsCommProtocal.SetBacklight(100)
        DelayMS delayTime
    End If

    If setTVInputSource = "HDMI" Then
        'Timing 69: HDMI-720P60
        'Timing 74: HDMI-1080P60
        Call frmCmbType.ChangeTiming("69")
        DelayMS delayTime
    End If

    If isAdjustOffset Then
        Label6.Caption = "GREY"

        Call frmCmbType.ChangePattern("110")
        DelayMS delayTime

        If isAdjustCool1 Then
            lbAdjustCOOL_1.BackColor = &H80FFFF
            Result = autoAdjustColorTemperature_Offset(valColorTempCool1, FixG, LowBri)
                
            If Result = False Then
                ShowError_Sys (11)
                GoTo FAIL
            Else
                'Call SAVE_WB_DATA_TO_ALL_SRC(setTVInputSource, setTVInputSourcePortNum)
                Call clsCommProtocal.SaveWBDataToAllSrc(setTVInputSource, setTVInputSourcePortNum)
                DelayMS delayTime
            End If
   
            lbAdjustCOOL_1.BackColor = &HC0FFC0
        End If
   
        If isAdjustNormal Then
            lbAdjustNormal.BackColor = &H80FFFF
            Result = autoAdjustColorTemperature_Offset(valColorTempNormal, FixG, LowBri)

            If Result = False Then
                ShowError_Sys (13)
                GoTo FAIL
            Else
                'Call SAVE_WB_DATA_TO_ALL_SRC(setTVInputSource, setTVInputSourcePortNum)
                Call clsCommProtocal.SaveWBDataToAllSrc(setTVInputSource, setTVInputSourcePortNum)
                DelayMS delayTime
            End If
    
            lbAdjustNormal.BackColor = &HC0FFC0
        End If
   
        If isAdjustWarm1 Then
            lbAdjustWARM_1.BackColor = &H80FFFF
            Result = autoAdjustColorTemperature_Offset(valColorTempWarm1, FixG, LowBri)
                
            If Result = False Then
                ShowError_Sys (14)
                GoTo FAIL
            Else
                'Call SAVE_WB_DATA_TO_ALL_SRC(setTVInputSource, setTVInputSourcePortNum)
                Call clsCommProtocal.SaveWBDataToAllSrc(setTVInputSource, setTVInputSourcePortNum)
                DelayMS delayTime
            End If

            lbAdjustWARM_1.BackColor = &HC0FFC0
        End If
    End If

    Call frmCmbType.ChangePattern("103")
    DelayMS delayTime
    
    Label6.Caption = "WHITE"

ADJUST_GAIN_AGAIN_COOL1:
    If isAdjustCool1 Then
        lbAdjustCOOL_1.BackColor = &H80FFFF
        Result = autoAdjustColorTemperature_Gain(valColorTempCool1, adjustMode3, HighBri)
  
        If Result = False Then
            ShowError_Sys (1)
            GoTo FAIL
        Else
            'Call SAVE_WB_DATA_TO_ALL_SRC(setTVInputSource, setTVInputSourcePortNum)
            Call clsCommProtocal.SaveWBDataToAllSrc(setTVInputSource, setTVInputSourcePortNum)
            DelayMS delayTime
        End If

        lbAdjustCOOL_1.BackColor = &HC0FFC0
        
        If adjustGainAgainCool1Flag > 0 Then
            GoTo CHECK_COOL1
        End If
    End If

ADJUST_GAIN_AGAIN_NORMAL:
    If isAdjustNormal Then
        lbAdjustNormal.BackColor = &H80FFFF
        Result = autoAdjustColorTemperature_Gain(valColorTempNormal, adjustMode3, HighBri)

        If Result = False Then
            ShowError_Sys (3)
            GoTo FAIL
        Else
            'Call SAVE_WB_DATA_TO_ALL_SRC(setTVInputSource, setTVInputSourcePortNum)
            Call clsCommProtocal.SaveWBDataToAllSrc(setTVInputSource, setTVInputSourcePortNum)
            DelayMS delayTime
        End If

        lbAdjustNormal.BackColor = &HC0FFC0
        
        If adjustGainAgainNormalFlag > 0 Then
            GoTo CHECK_NORMAL
        End If
    End If

ADJUST_GAIN_AGAIN_WARM1:
    If isAdjustWarm1 Then
        lbAdjustWARM_1.BackColor = &H80FFFF
        Result = autoAdjustColorTemperature_Gain(valColorTempWarm1, adjustMode3, HighBri)

        If Result = False Then
            ShowError_Sys (4)
            GoTo FAIL
        Else
            'Call SAVE_WB_DATA_TO_ALL_SRC(setTVInputSource, setTVInputSourcePortNum)
            Call clsCommProtocal.SaveWBDataToAllSrc(setTVInputSource, setTVInputSourcePortNum)
            DelayMS delayTime
        End If

        lbAdjustWARM_1.BackColor = &HC0FFC0
        
        If adjustGainAgainWarm1Flag > 0 Then
            GoTo CHECK_WARM1
        End If
    End If

    If isCheckColorTemp Then
CHECK_COOL1:
        If isAdjustCool1 Then
            Label6.Caption = "CHECK"
            lbAdjustCOOL_1.BackColor = &H80FFFF
            Result = checkColorAgain(valColorTempCool1, adjustMode3, HighBri)

            If Result = False Then
                ShowError_Sys (1)

                If adjustGainAgainCool1Flag > 0 Then
                    GoTo FAIL
                End If
                
                adjustGainAgainCool1Flag = adjustGainAgainCool1Flag + 1
                GoTo ADJUST_GAIN_AGAIN_COOL1
            End If
      
            lbAdjustCOOL_1.BackColor = &HC0FFC0
        End If
     
CHECK_NORMAL:
        If isAdjustNormal Then
            Label6.Caption = "CHECK"
            lbAdjustNormal.BackColor = &H80FFFF
            Result = checkColorAgain(valColorTempNormal, adjustMode3, HighBri)
      
            If Result = False Then
                ShowError_Sys (3)

                If adjustGainAgainNormalFlag > 0 Then
                    GoTo FAIL
                End If
    
                adjustGainAgainNormalFlag = adjustGainAgainNormalFlag + 1
                GoTo ADJUST_GAIN_AGAIN_NORMAL
            End If
    
            lbAdjustNormal.BackColor = &HC0FFC0
        End If

CHECK_WARM1:
        If isAdjustWarm1 Then
            Label6.Caption = "CHECK"
            lbAdjustWARM_1.BackColor = &H80FFFF
            Result = checkColorAgain(valColorTempWarm1, adjustMode3, HighBri)

            If Result = False Then
                ShowError_Sys (4)
                
                If adjustGainAgainWarm1Flag > 0 Then
                    GoTo FAIL
                End If
    
                adjustGainAgainWarm1Flag = adjustGainAgainWarm1Flag + 1
                GoTo ADJUST_GAIN_AGAIN_WARM1
            End If

            lbAdjustWARM_1.BackColor = &HC0FFC0
        End If
    End If
    
    'Last check:
    'Cool, 100% white pattern, brightness = 100, contrast = 100
    'Check Lv and save x, y, lv
    Call frmCmbType.ChangePattern("101")
    DelayMS delayTime
    
    'SET_BRIGHTNESS 100
    Call clsCommProtocal.SetBrightness(100)
    DelayMS delayTime
    Log_Info "Set brightness to 100"
    
    'SET_CONTRAST 100
    Call clsCommProtocal.SetContrast(100)
    DelayMS delayTime
    Log_Info "Set contrast to 100"
    
    'SET_BACKLIGHT 100
    Call clsCommProtocal.SetBacklight(100)
    DelayMS delayTime
    Log_Info "Set backlight to 100"
    
    'Call SET_COLORTEMP(valColorTempCool1, setTVInputSource, setTVInputSourcePortNum)
    Call clsCommProtocal.SetColorTemp(valColorTempCool1, setTVInputSource, setTVInputSourcePortNum)
    DelayMS delayTime
    Log_Info "Set color temp to cool1"

    DelayMS delayTime
    ObjCa.Measure
    rColorLastChk.xx = CLng(ObjProbe.sx * 10000)
    rColorLastChk.yy = CLng(ObjProbe.sy * 10000)
    rColorLastChk.lv = CLng(ObjProbe.lv)
    
    Log_Info "x = " + Str$(rColorLastChk.xx) + ", y = " + Str$(rColorLastChk.yy) + ", lv = " + Str$(rColorLastChk.lv)

    showData (lastChkShwDataStep)
    
    If rColorLastChk.lv < maxBrightnessSpec Then
        ShowError_Sys (30)
        GoTo FAIL
    End If

PASS:
    'EXIT_FAC_MODE
    clsCommProtocal.ExitFacMode
    DelayMS delayTime

    cmdMark = "PASS"
    Call saveALLcData

    CheckStep.ForeColor = &H80000008
    CheckStep.BackColor = &HC0FFC0
    CheckStep = CheckStep + strSerialNo + vbCrLf
    CheckStep = CheckStep + strSerialNo + vbCrLf
    CheckStep = CheckStep + "TEST ALL PASS"
    CheckStep.SelStart = Len(CheckStep)
    CheckStep.SetFocus
    checkResult.ForeColor = &HC000&
    checkResult.Caption = "PASS"
    DelayMS delayTime
    checkResult.BackColor = &HFF00&
    checkResult.ForeColor = &HC00000
    
    Label6.Caption = "PASS"
    
    Call subInitAfterRunning

    Exit Sub

FAIL:
    'EXIT_FAC_MODE
    clsCommProtocal.ExitFacMode
    DelayMS delayTime

    If IsSNWriteSuccess = funSNWrite Then
        cmdMark = "FAIL"
        Call saveALLcData
    End If

    CheckStep.SelStart = Len(CheckStep)
    CheckStep.SetFocus
    CheckStep.BackColor = &HFFFF&
    checkResult.BackColor = &HFF&
    checkResult.ForeColor = &H808080
    checkResult.Caption = "FAIL"
    DelayMS delayTime
    checkResult.ForeColor = &H0&
    checkResult.ForeColor = &HFFFF&
    
    Label6.Caption = "FAIL"

    Call subInitAfterRunning

    Exit Sub

ErrExit:
        MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Function funSNWrite() As Boolean
    strSerialNo = ""
    scanbarcode = ""
    strSerialNo = UCase$(txtInput.Text)

    If subJudgeTheSNIsAvailable = True Then
        funSNWrite = True
        scanbarcode = strSerialNo
    Else
        funSNWrite = False
    End If
End Function

Private Sub subInitBeforeRunning()
    countTime = 0
    lbTimer.Caption = "0s"
    Timer1.Enabled = True

    IsSNWriteSuccess = True
    txtInput.Locked = True
    strSerialNo = ""
    adjustGainAgainCool1Flag = 0
    adjustGainAgainNormalFlag = 0
    adjustGainAgainWarm1Flag = 0
End Sub

Private Sub subInitAfterRunning()
    Timer1.Enabled = False

    IsSNWriteSuccess = False
    adjustGainAgainCool1Flag = 0
    adjustGainAgainNormalFlag = 0
    adjustGainAgainWarm1Flag = 0

    txtInput.Text = ""
    txtInput.SetFocus
    txtInput.Locked = False
    
    If isUartMode = False Then
        isNetworkConnected = False
        tcpClient.Close
    End If
End Sub

Private Function subJudgeTheSNIsAvailable() As Boolean
    If strSerialNo = "" Or Len(strSerialNo) <> barCodeLen Then
        CheckStep.Text = ""
        CheckStep.Text = CheckStep.Text + "Please confirm the SN again?" + vbCrLf

        subJudgeTheSNIsAvailable = False
    Else
        subJudgeTheSNIsAvailable = True
        
        Set cn = Nothing
        Set rs = Nothing
        sqlstring = ""
    End If
End Function

Sub ShowError_Sys(t As Integer)
    Dim s As String
    
    s = "Unknown"

    Select Case t
        Case 1
            s = "ColorTemp_COOL_1 is Wrong, Please Check Again."
        Case 2
            s = "ColorTemp_COOL_2 is Wrong, Please Check Again."
        Case 3
            s = "ColorTemp_NORMAL is Wrong, Please Check Again."
        Case 4
            s = "ColorTemp_WARM_1 is Wrong, Please Check Again."
        Case 5
            s = "ColorTemp_WARM_2 is Wrong, Please Check Again."
        Case 6
            s = "LAB_SN:" + strSerialNo + "(End)  Len:" + Str$(barCodeLen) + vbCrLf + "条形码长度不对，请确认！"
        Case 7
            s = "Can not Write DVI EDID."
        Case 8
            s = "Calibrate FAIL.(AUTO LEVEL)"
        Case 9
            s = "RS232 Connector Error"
        Case 10
            s = "Read DSUB EDID FAIL"
        Case 11
            s = "OFFSET_Color_COOL_1 is Wrong, Please Check Again."
        Case 12
            s = "OFFSET_Color_COOL_2 is Wrong, Please Check Again."
        Case 13
            s = "OFFSET_Color_NORMAL is Wrong, Please Check Again."
        Case 14
            s = "OFFSET_Color_WARM_1 is Wrong, Please Check Again."
        Case 15
            s = "OFFSET_Color_WARM_2 is Wrong, Please Check Again."
        Case 16
            s = "HDMI2 CheckSum is Wrong"
        Case 17
            s = "Can not Write HDMI-2 EDID."
        Case 18
            s = "min_Brightness is over SPEC."
        Case 19
            s = "FW Version is Wrong."
        Case 20
            s = "Can not Write OSD-SN."
        Case 21
            s = "max_Brightness is over SPEC."
        Case 22
            s = "ColorTemp_COOL_1 is Wrong, Please Check Again."
        Case 23
            s = "ColorTemp_COOL_2 is Wrong, Please Check Again."
        Case 24
            s = "ColorTemp_NORMAL is Wrong, Please Check Again."
        Case 25
            s = "ColorTemp_WARM_1 is Wrong, Please Check Again."
        Case 26
            s = "ColorTemp_WARM_2 is Wrong, Please Check Again."
        Case 27
            s = "ColorTemp_5000 is Wrong, Please Check Again."
        Case 28
            s = "ColorTemp_3000 is Wrong, Please Check Again."
        Case 29
            s = "LightSensor Data is Wrong, Please Check Again."
        Case 30
            s = "亮度不在规格！"
        Case 31
            s = ""
    End Select

    CheckStep.ForeColor = &HFF&
    CheckStep.Text = CheckStep.Text + "Error Code:" + Str$(t) + vbCrLf + s + vbCrLf
    CheckStep.SelStart = Len(CheckStep)
    CheckStep.SetFocus
End Sub

Private Function autoAdjustColorTemperature_Gain(ColorTemp As Long, adjustVal As Long, HighLowMode As Long) As Boolean
    Dim i, j, k As Integer
  
    Log_Info "========Adjust " + Str$(ColorTemp) + "K========"
  
    For j = 1 To 2
        'Call SET_COLORTEMP(ColorTemp, setTVInputSource, setTVInputSourcePortNum)
        Call clsCommProtocal.SetColorTemp(ColorTemp, setTVInputSource, setTVInputSourcePortNum)
        DelayMS delayTime
  
        Call setColorTemp(ColorTemp, presetData, HighLowMode)
        DelayMS delayTime
        
        Log_Info "Init current colorTemp. RES:" + Str$(RES)
        rRGB.cRR = presetData.nColorRR
        rRGB.cGG = presetData.nColorGG
        rRGB.cBB = presetData.nColorBB
        
        Label1 = Str$(presetData.xx)
        Label3 = Str$(presetData.yy)

        'SET_R_GAN rRGB.cRR
        Call clsCommProtocal.SetRGain(rRGB.cRR)
        DelayMS delayTime
        
        'SET_G_GAN rRGB.cGG
        Call clsCommProtocal.SetGGain(rRGB.cGG)
        DelayMS delayTime
        
        'SET_B_GAN rRGB.cBB
        Call clsCommProtocal.SetBGain(rRGB.cBB)
        DelayMS delayTime

        showData (1)

        resCodeForAdjustColorTemp = 0
        
        For k = 1 To 50
            If IsStop = True Then GoTo Cancel
            
            RES = checkColorTemp(rColor, ColorTemp)
            Log_Info "Check colorTemp. RES:" + Str$(RES)
            
            If RES Then Exit For
            
            If RES = False Then
                If resCodeForAdjustColorTemp = 0 Then
                    Call adjustColorTemp(adjustMode3, AdjustSingle, SingleStep, rRGB, resCodeForAdjustColorTemp)
                ElseIf resCodeForAdjustColorTemp = 1 Then
                    Call adjustColorTemp(adjustMode1, AdjustSingle, SingleStep, rRGB, resCodeForAdjustColorTemp)
                ElseIf resCodeForAdjustColorTemp = 2 Then
                    Call adjustColorTemp(adjustMode2, AdjustSingle, SingleStep, rRGB, resCodeForAdjustColorTemp)
                ElseIf resCodeForAdjustColorTemp = 3 Then
                    Call adjustColorTemp(adjustMode4, AdjustSingle, SingleStep, rRGB, resCodeForAdjustColorTemp)
                End If
                Log_Info "SET_RGB_GAN: R = " + Str$(rRGB.cRR) + ", G = " + Str$(rRGB.cGG) + ", B = " + Str$(rRGB.cBB) + ", resultcode = " + Str$(resCodeForAdjustColorTemp)
 
                'SET_R_GAN rRGB.cRR
                Call clsCommProtocal.SetRGain(rRGB.cRR)
                DelayMS delayTime
                
                'SET_G_GAN rRGB.cGG
                Call clsCommProtocal.SetGGain(rRGB.cGG)
                DelayMS delayTime
                
                'SET_B_GAN rRGB.cBB
                Call clsCommProtocal.SetBGain(rRGB.cBB)
                DelayMS delayTime

                showData (2)
            End If
        Next k
  
        If isAdjustOffset Then
            Call LoadData(ColorTemp, False)
            
            'SET_R_OFF rRGB1.cRR
            Call clsCommProtocal.SetROffset(rRGB1.cRR)
            DelayMS delayTime
               
            'SET_G_OFF rRGB1.cGG
            Call clsCommProtocal.SetGOffset(rRGB1.cGG)
            DelayMS delayTime
            
            'SET_B_OFF rRGB1.cBB
            Call clsCommProtocal.SetBOffset(rRGB1.cBB)
            DelayMS delayTime
        Else
            'SET_R_OFF 128
            Call clsCommProtocal.SetROffset(128)
            DelayMS delayTime
           
            'SET_G_OFF 128
            Call clsCommProtocal.SetGOffset(128)
            DelayMS delayTime
        
            'SET_B_OFF 128
            Call clsCommProtocal.SetBOffset(128)
            DelayMS delayTime
        End If
        
        If RES Then Exit For
        
    Next j

Cancel:
    If RES Then
        Call saveData(ColorTemp, HighLowMode)
        Log_Info "Save current data of colorTemp."
        autoAdjustColorTemperature_Gain = True
    Else
        autoAdjustColorTemperature_Gain = False
    End If

End Function

Private Function autoAdjustColorTemperature_Offset(ColorTemp As Long, FixValue As Long, HighLowMode As Long) As Boolean
    Dim i, j, k As Integer
  
    Log_Info "========Adjust " + Str$(ColorTemp) + "K========"
  
    For j = 1 To 2
        'Call SET_COLORTEMP(ColorTemp, setTVInputSource, setTVInputSourcePortNum)
        Call clsCommProtocal.SetColorTemp(ColorTemp, setTVInputSource, setTVInputSourcePortNum)
        DelayMS delayTime

        Call setColorTemp(ColorTemp, presetData, HighLowMode)
        DelayMS delayTime
        Log_Info "Init current colorTemp. RES:" + Str$(RES)
        rRGB.cRR = presetData.nColorRR
        rRGB.cGG = presetData.nColorGG
        rRGB.cBB = presetData.nColorBB
  
        'Label1 = Str$(presetData.xx)
        'Label3 = Str$(presetData.yy)

        'Call LoadData(ColorTemp, True)

        'SET_R_GAN rRGB1.cRR
        'DelayMs StepTime
        
        'SET_G_GAN rRGB1.cGG
        'DelayMs StepTime

        'SET_B_GAN rRGB1.cBB
        'DelayMs StepTime

        'SET_R_OFF rRGB.cRR
        Call clsCommProtocal.SetROffset(rRGB.cRR)
        DelayMS delayTime

        'SET_G_OFF rRGB.cGG
        Call clsCommProtocal.SetGOffset(rRGB.cGG)
        DelayMS delayTime
        
        'SET_B_OFF rRGB.cBB
        Call clsCommProtocal.SetBOffset(rRGB.cBB)
        DelayMS delayTime

        If False Then
            showData (3)

            For k = 1 To 50
                If IsStop = True Then GoTo Cancel
                
                RES = checkColorTempOffset(rColor, ColorTemp)
                Log_Info "Check colorTemp. RES:" + Str$(RES)
    
                If RES Then Exit For
                If RES = False Then
                    Call adjustColorTempOffset(FixValue, AdjustSingle, SingleStep, rRGB)
    
                    'SET_R_OFF rRGB.cRR
                    Call clsCommProtocal.SetROffset(rRGB.cRR)
                    DelayMS delayTime
                    
                    'SET_G_OFF rRGB.cGG
                    Call clsCommProtocal.SetGOffset(rRGB.cGG)
                    DelayMS delayTime
    
                    'SET_B_OFF rRGB.cBB
                    Call clsCommProtocal.SetBOffset(rRGB.cBB)
                    DelayMS delayTime
    
                    showData (4)
                End If
    
                DelayMS 200
            Next k
        Else
            Call saveData(ColorTemp, HighLowMode)
            autoAdjustColorTemperature_Offset = True
            
            Exit Function
        End If

        If RES Then Exit For

        DelayMS delayTime
    Next j
 
Cancel:
    If RES Then
        Call saveData(ColorTemp, HighLowMode)
        Log_Info "Save current data of colorTemp."
        autoAdjustColorTemperature_Offset = True
    Else
        autoAdjustColorTemperature_Offset = False
    End If

End Function

Private Function checkColorAgain(ColorTemp As Long, adjustVal As Long, HighLowMode As Long) As Boolean
    Dim i, j, k As Integer
  
    Log_Info "========Check " + Str$(ColorTemp) + "K========"
  
    For j = 1 To 2
        'Call SET_COLORTEMP(ColorTemp, setTVInputSource, setTVInputSourcePortNum)
        Call clsCommProtocal.SetColorTemp(ColorTemp, setTVInputSource, setTVInputSourcePortNum)
  
        Call setColorTemp(ColorTemp, presetData, HighLowMode)
        DelayMS delayTime
        Log_Info "Init current colorTemp. RES:" + Str$(RES)

        Label1 = Str$(presetData.xx)
        Label3 = Str$(presetData.yy)

        showData (5)

        If IsStop = True Then GoTo Cancel

        RES = checkColorTemp(rColor, ColorTemp)
        Log_Info "Check colorTemp. RES:" + Str$(RES)

        If RES Then Exit For

        DelayMS delayTime
    Next j
  
Cancel:
    If RES Then
        checkColorAgain = True
    Else
        checkColorAgain = False
    End If

End Function


'step = lastChkShwDataStep: Check max brightness of TV with brightness 100 and contrast 100 in 100% white pattern.
Private Sub showData(step As Integer)
On Error Resume Next
    Dim xPos, yPos, vPos As Long
    
    DelayMS delayTime
    ObjCa.Measure
    rColor.xx = CLng(ObjProbe.sx * 10000)
    rColor.yy = CLng(ObjProbe.sy * 10000)
    rColor.lv = CLng(ObjProbe.lv)
  
    Picture1.Cls
    
    'The values here are about 15 times bigger than the actual pixel.
    '(1515,1275) is the origin of dx-dy axis.
    'In lv axis, 1660 is the distance from the bottom edge of blue rectangle to the top of Picture1.
    'In dx, 365 is half a side of blue rectangle.
    xPos = 1515 + (rColor.xx - presetData.xx) * 365 / presetData.xt
    yPos = 1275 - (rColor.yy - presetData.yy) * 385 / presetData.yt
    If step = lastChkShwDataStep Then
        vPos = 1660 - (rColor.lv - maxBrightnessSpec) * 385 / 50
    Else
        vPos = 1660 - (rColor.lv - presetData.lv) * 385 / 50
    End If

    'In dx-dy axis, 360 is the distance from left edge of white rectangle to the left of Picture1.
    'In dx-dy axis, 2660 is the distance from right edge of white rectangle to the left of Picture1.
    'In dx-dy axis, 80 is the distance from top edge of white rectangle to the top of Picture1.
    'In dx-dy axis, 2660 is the distance from bottom edge of white rectangle to the top of Picture1.
    If xPos < 360 Then xPos = 360
    If xPos > 2660 Then xPos = 2660
    If yPos < 80 Then yPos = 80
    If yPos > 2480 Then yPos = 2480

    If step <> lastChkShwDataStep Then
        If Abs(rColor.xx - presetData.xx) <= presetData.xt And Abs(rColor.yy - presetData.yy) <= presetData.yt Then
            lbColorTempWrong.Visible = False
            Picture1.Circle (xPos, yPos), 23, &H30FF30
        Else
            lbColorTempWrong.Visible = True
            Picture1.Circle (xPos, yPos), 23, &HFF&
          
            If rColor.xx < 5 Then
                IsStop = True
                ObjCa.RemoteMode = 2
                MsgBox ("Please check the CA210 Probe is OK or not.")
                RES = False
            End If
        End If
    End If

    'In lv axis, 3060 is the distance from left edge of white rectangle to the left of Picture1.
    'In lv axis, 3390 is the distance from right edge of white rectangle to the left of Picture1.
    If step = lastChkShwDataStep Then
        If rColor.lv > maxBrightnessSpec Then
            Picture1.Line (3060, vPos)-(3390, vPos), &H30FF30
        Else
            Picture1.Line (3060, vPos)-(3390, vPos), &HFF&
        End If
    Else
        If rColor.lv > presetData.lv Then
            Picture1.Line (3060, vPos)-(3390, vPos), &H30FF30
        Else
            Picture1.Line (3060, vPos)-(3390, vPos), &HFF&
        End If
    End If
 
    Log_Info "_x/y/Lv: " + Str$(rColor.xx) + " / " + Str$(rColor.yy) + " / " + Str$(rColor.lv)

    If Label6 <> "CHECK" Then Log_Info "_R/G/B:  " + Str$(rRGB.cRR) + " / " + Str$(rRGB.cGG) + " / " + Str$(rRGB.cBB)

    Label_x = Str$(rColor.xx)
    Label_y = Str$(rColor.yy)
    Label_Lv = Str$(rColor.lv)
    DelayMS 30
End Sub

Private Sub tbDisConnectastro_Click()
    ObjCa.RemoteMode = 0
End Sub

Private Sub Timer1_Timer()
    countTime = countTime + 1
    lbTimer.Caption = CStr(countTime) & "s"
End Sub

Private Sub vbConChroma_Click()
    frmCmbType.Show
End Sub

Private Sub vbSetSPEC_Click()
    frmSetData.Show
End Sub

Private Sub vbAbout_Click()
    frmAbout.Show
End Sub

Private Sub vbConCA310_Click()
    If IsCa210ok = True Then
        ObjCa.RemoteMode = 1
        Exit Sub
    Else
        CONNECT_CA210
    End If
End Sub


Private Sub Form_Load()
    i = 0
    IsStop = False
    txtInput.Locked = False
    
    gstrBrand = Split(gstrCurProjName, gstrDelimiterForProjName)(0)
    
    If gstrBrand = "Letv" Then
        Set clsLetvProtocal = New LetvProtocal
        Set clsCommProtocal = clsLetvProtocal
    ElseIf gstrBrand = "Haier" Or gstrBrand = "CIBN" Then
        Set clsHaierProtocal = New HaierProtocal
        Set clsCommProtocal = clsHaierProtocal
    Else
        Set clsLetvProtocal = New LetvProtocal
        Set clsCommProtocal = clsLetvProtocal
    End If

    subInitInterface
    
    RES = initColorTemp(Calibrate, MinBrightness, gstrCurProjName, App.Path)      'InitLPT in dll.
End Sub

Public Sub subInitInterface()
    Dim clsConfigData As ProjectConfig

    Set clsConfigData = New ProjectConfig
    clsConfigData.LoadConfigData
    
    setTVCurrentComBaud = clsConfigData.ComBaud
    setTVCurrentComID = clsConfigData.ComID
    setTVInputSource = clsConfigData.inputSource
    setTVInputSourcePortNum = CInt(Right(setTVInputSource, 1))
    setTVInputSource = Left(setTVInputSource, Len(setTVInputSource) - 1)
    delayTime = clsConfigData.DelayMS
    Ca210ChannelNO = clsConfigData.ChannelNum
    barCodeLen = clsConfigData.barCodeLen
    maxBrightnessSpec = clsConfigData.LvSpec
    isAdjustCool2 = clsConfigData.EnableCool2
    isAdjustCool1 = clsConfigData.EnableCool1
    isAdjustNormal = clsConfigData.EnableNormal
    isAdjustWarm1 = clsConfigData.EnableWarm1
    isAdjustWarm2 = clsConfigData.EnableWarm2
    isCheckColorTemp = clsConfigData.EnableChkColor
    isAdjustOffset = clsConfigData.EnableAdjOffset
    
    If clsConfigData.CommMode = modeUART Then
        isUartMode = True
        lbCommMode.Caption = "UART"
        subInitComPort
    Else
        isUartMode = False
        lbCommMode.Caption = "Network"
        subInitNetwork
    End If
    
    Set clsConfigData = Nothing

    txtInput.Text = ""
    lbModelName.Caption = gstrCurProjName
    
    If isAdjustCool1 = True Then lbAdjustCOOL_1.ForeColor = &H80000008
    If isAdjustCool2 = True Then lbAdjustCOOL_2.ForeColor = &H80000008
    If isAdjustNormal = True Then lbAdjustNormal.ForeColor = &H80000008
    If isAdjustWarm1 = True Then lbAdjustWARM_1.ForeColor = &H80000008
    If isAdjustWarm2 = True Then lbAdjustWARM_2.ForeColor = &H80000008

    If isAdjustCool1 = False Then lbAdjustCOOL_1.ForeColor = &HC0C0C0
    If isAdjustCool2 = False Then lbAdjustCOOL_2.ForeColor = &HC0C0C0
    If isAdjustNormal = False Then lbAdjustNormal.ForeColor = &HC0C0C0
    If isAdjustWarm1 = False Then lbAdjustWARM_1.ForeColor = &HC0C0C0
    If isAdjustWarm2 = False Then lbAdjustWARM_2.ForeColor = &HC0C0C0
End Sub

Private Sub subInitComPort()
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
    
    MSComm1.CommPort = setTVCurrentComID
    MSComm1.Settings = setTVCurrentComBaud & ",N,8,1"
    MSComm1.InputLen = 0
        
    MSComm1.InBufferCount = 0
    MSComm1.OutBufferCount = 0
    MSComm1.InputMode = comInputModeBinary
        
    MSComm1.NullDiscard = False
    MSComm1.DTREnable = False
    MSComm1.EOFEnable = False
    MSComm1.RTSEnable = False
    MSComm1.SThreshold = 1
    MSComm1.RThreshold = 1
    MSComm1.InBufferSize = 1024
    MSComm1.OutBufferSize = 512
End Sub

Private Sub subInitNetwork()
    isNetworkConnected = False
    With tcpClient
        .Protocol = sckTCPProtocol
        ' IMPORTANT: be sure to change the RemoteHost
        ' value to the name of your computer.
        .RemoteHost = strRemoteHost
        .RemotePort = lngRemotePort
    End With
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
On Error GoTo ErrExit

    If KeyAscii = 13 Then
        IsStop = False
        
        If txtInput.Locked = False Then
            If isUartMode = True Then
                MSComm1.PortOpen = True
                subMainProcesser
            Else
                isNetworkConnected = False
                Do
                    Log_Info "TCP Connect"
                    tcpClient.Connect
                    Call DelaySWithFlag(cmdReceiveWaitS * 2, isNetworkConnected)
                
                    If isNetworkConnected = True Then
                        subMainProcesser
                        Exit Do
                    Else
                        tcpClient.Close
                        i = i + 1
                    End If
                    Log_Info "Re-connect to TV."
                Loop While i <= 5
            End If
        End If
        
        If IsStop = True Then
            Exit Sub
        End If
    End If
    Exit Sub

ErrExit:
    'Invalid Port Number
    If Err.Number = 8002 Then
        MsgBox Err.Description, vbCritical, Err.Source
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrExit

    If gstrBrand = "Letv" Then
        Set clsLetvProtocal = Nothing
        Set clsCommProtocal = Nothing
    ElseIf gstrBrand = "Haier" Or gstrBrand = "CIBN" Then
        Set clsHaierProtocal = Nothing
        Set clsCommProtocal = Nothing
    Else
        Set clsLetvProtocal = Nothing
        Set clsCommProtocal = Nothing
    End If

    IsStop = True
    If (IsCa210ok = True) Then
        ObjCa.RemoteMode = 0
    End If
  
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
  
    Call DeinitColorTemp(gstrCurProjName)
    End
    Exit Sub

ErrExit:
        MsgBox Err.Description, vbCritical, Err.Source
End Sub


Private Sub saveData(ColorTemp As Long, HL As Long)

    Select Case ColorTemp
        Case valColorTempCool1
            If HL Then
                c12000K.xx = rColor.xx
                c12000K.yy = rColor.yy
                c12000K.lv = rColor.lv
                c12000K.nColorRR = rRGB.cRR
                c12000K.nColorGG = rRGB.cGG
                c12000K.nColorBB = rRGB.cBB
            Else
                cFF12000K.xx = rColor.xx
                cFF12000K.yy = rColor.yy
                cFF12000K.lv = rColor.lv
                cFF12000K.nColorRR = rRGB.cRR
                cFF12000K.nColorGG = rRGB.cGG
                cFF12000K.nColorBB = rRGB.cBB
            End If

        Case valColorTempNormal
            If HL Then
                c10000K.xx = rColor.xx
                c10000K.yy = rColor.yy
                c10000K.lv = rColor.lv
                c10000K.nColorRR = rRGB.cRR
                c10000K.nColorGG = rRGB.cGG
                c10000K.nColorBB = rRGB.cBB
            Else
                cFF10000K.xx = rColor.xx
                cFF10000K.yy = rColor.yy
                cFF10000K.lv = rColor.lv
                cFF10000K.nColorRR = rRGB.cRR
                cFF10000K.nColorGG = rRGB.cGG
                cFF10000K.nColorBB = rRGB.cBB
            End If

        Case valColorTempWarm1
            If HL Then
                c6500K.xx = rColor.xx
                c6500K.yy = rColor.yy
                c6500K.lv = rColor.lv
                c6500K.nColorRR = rRGB.cRR
                c6500K.nColorGG = rRGB.cGG
                c6500K.nColorBB = rRGB.cBB
            Else
                cFF6500K.xx = rColor.xx
                cFF6500K.yy = rColor.yy
                cFF6500K.lv = rColor.lv
                cFF6500K.nColorRR = rRGB.cRR
                cFF6500K.nColorGG = rRGB.cGG
                cFF6500K.nColorBB = rRGB.cBB
            End If
    End Select
  
End Sub

Private Sub LoadData(ColorTemp As Long, isGain As Boolean)
    Select Case ColorTemp
        Case valColorTempCool1
            If isGain Then
                rRGB1.cRR = c12000K.nColorRR
                rRGB1.cGG = c12000K.nColorGG
                rRGB1.cBB = c12000K.nColorBB
            Else
                rRGB1.cRR = cFF12000K.nColorRR
                rRGB1.cGG = cFF12000K.nColorGG
                rRGB1.cBB = cFF12000K.nColorBB
            End If
            
        Case valColorTempNormal
            If isGain Then
                rRGB1.cRR = c10000K.nColorRR
                rRGB1.cGG = c10000K.nColorGG
                rRGB1.cBB = c10000K.nColorBB
            Else
                rRGB1.cRR = cFF10000K.nColorRR
                rRGB1.cGG = cFF10000K.nColorGG
                rRGB1.cBB = cFF10000K.nColorBB
            End If
            
        Case valColorTempWarm1
            If isGain Then
                rRGB1.cRR = c6500K.nColorRR
                rRGB1.cGG = c6500K.nColorGG
                rRGB1.cBB = c6500K.nColorBB
            Else
                rRGB1.cRR = cFF6500K.nColorRR
                rRGB1.cGG = cFF6500K.nColorGG
                rRGB1.cBB = cFF6500K.nColorBB
            End If
    End Select
End Sub

Private Sub saveALLcData()
    If strSerialNo = "" Then
        Exit Sub
    Else
        sqlstring = "select * from DataRecord"
        Executesql (sqlstring)
        rs.AddNew

        rs.Fields(0) = gstrCurProjName
        rs.Fields(1) = strSerialNo

        rs.Fields(2) = c12000K.xx
        rs.Fields(3) = c12000K.yy
        rs.Fields(4) = c12000K.nColorRR
        rs.Fields(5) = c12000K.nColorGG
        rs.Fields(6) = c12000K.nColorBB
        rs.Fields(7) = c10000K.xx
        rs.Fields(8) = c10000K.yy
        rs.Fields(9) = c10000K.nColorRR
        rs.Fields(10) = c10000K.nColorGG
        rs.Fields(11) = c10000K.nColorBB
        rs.Fields(12) = c6500K.xx
        rs.Fields(13) = c6500K.yy
        rs.Fields(14) = c6500K.nColorRR
        rs.Fields(15) = c6500K.nColorGG
        rs.Fields(16) = c6500K.nColorBB
        
        rs.Fields(17) = cFF12000K.nColorRR
        rs.Fields(18) = cFF12000K.nColorGG
        rs.Fields(19) = cFF12000K.nColorBB
        rs.Fields(20) = cFF10000K.nColorRR
        rs.Fields(21) = cFF10000K.nColorGG
        rs.Fields(22) = cFF10000K.nColorBB
        rs.Fields(23) = cFF6500K.nColorRR
        rs.Fields(24) = cFF6500K.nColorGG
        rs.Fields(25) = cFF6500K.nColorBB

        rs.Fields(26) = rColorLastChk.lv
        rs.Fields(27) = maxBrightnessSpec

        rs.Fields(28) = cmdMark
        rs.Fields(29) = Date
        rs.Fields(30) = Time

        rs.Update

        Set cn = Nothing
        Set rs = Nothing
        sqlstring = ""
    End If
End Sub

Private Sub tcpClient_Connect()
    'Success to connect the TV.
    isNetworkConnected = True
End Sub
