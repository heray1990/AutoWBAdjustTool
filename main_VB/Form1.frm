VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto White Balance Modulation"
   ClientHeight    =   4620
   ClientLeft      =   5865
   ClientTop       =   2625
   ClientWidth     =   10230
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
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PictureBrand 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   758
      Left            =   120
      Picture         =   "Form1.frx":1DF72
      ScaleHeight     =   735
      ScaleWidth      =   2505
      TabIndex        =   21
      Top             =   0
      Width           =   2528
   End
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
      Picture         =   "Form1.frx":24226
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
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3580
      Left            =   6440
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Form1.frx":43320
      Top             =   960
      Width           =   3700
   End
   Begin VB.TextBox txtInput 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Text            =   "123456789"
      Top             =   1130
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "----"
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
      Left            =   4560
      TabIndex        =   20
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label lbModelName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sampl1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   120
      TabIndex        =   19
      Top             =   750
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
      Alignment       =   2  'Center
      Caption         =   "----"
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
      Width           =   900
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
      Width           =   3810
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
      Caption         =   "----"
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
      Left            =   5490
      TabIndex        =   12
      Top             =   3525
      Width           =   960
   End
   Begin VB.Label Label_y 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "----"
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
      Height          =   405
      Left            =   4560
      TabIndex        =   11
      Top             =   3555
      Width           =   900
   End
   Begin VB.Label Label_x 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "----"
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
      Height          =   405
      Left            =   3120
      TabIndex        =   10
      Top             =   3555
      Width           =   900
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
      Width           =   7500
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " x:"
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
      Width           =   1440
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
      Left            =   4070
      TabIndex        =   9
      Top             =   3525
      Width           =   1440
   End
   Begin VB.Menu vbFunc 
      Caption         =   "Function"
      Begin VB.Menu vbConCA310 
         Caption         =   "Connect CA310/CA210"
      End
      Begin VB.Menu tbDisConnectastro 
         Caption         =   "DisConnect CA310/CA210(&D)"
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
Dim cCOOL1 As COLORTEMPSPEC
Dim cNORMAL As COLORTEMPSPEC
Dim cWARM1 As COLORTEMPSPEC
Dim cFFCOOL1 As COLORTEMPSPEC
Dim cFFNORMAL As COLORTEMPSPEC
Dim cFFWARM1 As COLORTEMPSPEC
Dim rColor As REALCOLOR
Dim lvLastChk As Long
Dim Calibrate, MinBrightness As Long
Dim resCodeForAdjustColorTemp As Long
Dim cmdMark As String
Dim clsProtocal As Protocal
Dim clsCANTVProtocal As CANTVProtocal
Dim clsLetvProtocal As LetvProtocal
Dim clsLetvCurvedProtocal As LetvCurvedProtocal
Dim clsLetvMST6M60 As LetvMST6M60
Dim clsHaierProtocal As HaierProtocal

Dim ivpg As IVPGCtrl
Private mTitle As String
Dim m_Title As String

Private WithEvents Obj As VPGCtrl.VPGCtrl
Attribute Obj.VB_VarHelpID = -1

Private Sub subMainProcesser()
    Dim i, j As Integer

On Error GoTo ErrExit
    subInitBeforeRunning

    If IsStop = True Then
        Exit Sub
    End If

    If IsCa210ok = False Then
        MsgBox TXTCaDisconnectHint, vbOKOnly + vbInformation, "warning"
        subInitAfterRunning
        
        Exit Sub
    End If

    checkResult.BackColor = &H80FFFF
    IsStop = False
    checkResult.Caption = "RUN..."
    checkResult.ForeColor = &HC0&
    CheckStep = ""

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

    Log_Info "Start adjusting color temperature"
    Call ChangePattern(gstrVPG80IRE)

    clsProtocal.EnterFacMode
    Call clsProtocal.SwitchInputSource(setTVInputSource, setTVInputSourcePortNum)
    Call clsProtocal.ResetPicMode
    Call clsProtocal.SetBacklight(100)
    Log_Info "Set backlight to 100"

    Label6.Caption = "WHITE"

ADJUST_GAIN_AGAIN_COOL1:
    If isAdjustCool1 Then
        lbAdjustCOOL_1.BackColor = &H80FFFF
        Result = AdjRGBGain(cstrColorTempCool1, adjustMode3, HighBri)
  
        If Result = False Then
            ShowError_Sys (1)
            GoTo FAIL
        Else
            Call clsProtocal.SaveWBDataToAllSrc(setTVInputSource, setTVInputSourcePortNum)
        End If

        SaveLogInFile "[Time]White Cool1: " & lbTimer.Caption
        lbAdjustCOOL_1.BackColor = &HC0FFC0
        
        If adjustGainAgainCool1Flag > 0 Then
            GoTo CHECK_COOL1
        End If
    End If

ADJUST_GAIN_AGAIN_NORMAL:
    If isAdjustNormal Then
        lbAdjustNormal.BackColor = &H80FFFF
        Result = AdjRGBGain(cstrColorTempNormal, adjustMode3, HighBri)

        If Result = False Then
            ShowError_Sys (3)
            GoTo FAIL
        Else
            Call clsProtocal.SaveWBDataToAllSrc(setTVInputSource, setTVInputSourcePortNum)
        End If

        SaveLogInFile "[Time]White Normal: " & lbTimer.Caption
        lbAdjustNormal.BackColor = &HC0FFC0
        
        If adjustGainAgainNormalFlag > 0 Then
            GoTo CHECK_NORMAL
        End If
    End If

ADJUST_GAIN_AGAIN_WARM1:
    If isAdjustWarm1 Then
        lbAdjustWARM_1.BackColor = &H80FFFF
        Result = AdjRGBGain(cstrColorTempWarm1, adjustMode3, HighBri)

        If Result = False Then
            ShowError_Sys (4)
            GoTo FAIL
        Else
            Call clsProtocal.SaveWBDataToAllSrc(setTVInputSource, setTVInputSourcePortNum)
        End If

        SaveLogInFile "[Time]White Warm1: " & lbTimer.Caption
        lbAdjustWARM_1.BackColor = &HC0FFC0
        
        If adjustGainAgainWarm1Flag > 0 Then
            GoTo CHECK_WARM1
        End If
    End If

    If isAdjustOffset Then
        Label6.Caption = "GREY"

        Call ChangePattern(gstrVPG20IRE)

        If isAdjustCool1 Then
            lbAdjustCOOL_1.BackColor = &H80FFFF
            Result = AdjRGBOffset(cstrColorTempCool1, FixG, LowBri)
                
            If Result = False Then
                ShowError_Sys (11)
                GoTo FAIL
            Else
                Call clsProtocal.SaveWBDataToAllSrc(setTVInputSource, setTVInputSourcePortNum)
            End If
            
            SaveLogInFile "[Time]Grey Cool1: " & lbTimer.Caption
            lbAdjustCOOL_1.BackColor = &HC0FFC0
        End If
   
        If isAdjustNormal Then
            lbAdjustNormal.BackColor = &H80FFFF
            Result = AdjRGBOffset(cstrColorTempNormal, FixG, LowBri)

            If Result = False Then
                ShowError_Sys (13)
                GoTo FAIL
            Else
                Call clsProtocal.SaveWBDataToAllSrc(setTVInputSource, setTVInputSourcePortNum)
            End If

            SaveLogInFile "[Time]Grey Normal: " & lbTimer.Caption
            lbAdjustNormal.BackColor = &HC0FFC0
        End If
   
        If isAdjustWarm1 Then
            lbAdjustWARM_1.BackColor = &H80FFFF
            Result = AdjRGBOffset(cstrColorTempWarm1, FixG, LowBri)
                
            If Result = False Then
                ShowError_Sys (14)
                GoTo FAIL
            Else
                Call clsProtocal.SaveWBDataToAllSrc(setTVInputSource, setTVInputSourcePortNum)
            End If

            SaveLogInFile "[Time]Grey Warm1: " & lbTimer.Caption
            lbAdjustWARM_1.BackColor = &HC0FFC0
        End If
    End If

    If isCheckColorTemp Then
        If isAdjustOffset Then
            Call ChangePattern(gstrVPG80IRE)
        End If

CHECK_COOL1:
        If isAdjustCool1 Then
            Label6.Caption = "CHECK"
            lbAdjustCOOL_1.BackColor = &H80FFFF
            Result = checkColorAgain(cstrColorTempCool1, HighBri)

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
            Result = checkColorAgain(cstrColorTempNormal, HighBri)

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
            Result = checkColorAgain(cstrColorTempWarm1, HighBri)

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
    
    If gstrChipSet = "T111" Then
        Call clsProtocal.SelColorTemp(cstrColorTempNormal, setTVInputSource, setTVInputSourcePortNum)
        Log_Info "Set color temp to cool1"
        
        ObjCa.Measure
        lvLastChk = CLng(ObjProbe.lv)
        Log_Info "lv = " + CStr(lvLastChk)
        showData (lastChkShwDataStep)
        
        If lvLastChk <= maxBrightnessSpec Then
            ShowError_Sys (30)
            GoTo FAIL
        End If
    Else
        'Last check:
        'Cool, 100% white pattern, brightness = 100, contrast = 100
        'Check Lv and save x, y, lv
        Call ChangePattern(gstrVPG100IRE)

        Call clsProtocal.SetBrightness(100)
        Log_Info "Set brightness to 100"

        Call clsProtocal.SetContrast(100)
        Log_Info "Set contrast to 100"

        Call clsProtocal.SelColorTemp(cstrColorTempCool1, setTVInputSource, setTVInputSourcePortNum)
        Log_Info "Set color temp to cool1"

        ObjCa.Measure
        lvLastChk = CLng(ObjProbe.lv)
        Log_Info "lv = " + CStr(lvLastChk)
        showData (lastChkShwDataStep)

        Call clsProtocal.SetBrightness(50)
        Call clsProtocal.SetContrast(50)
        Log_Info "Set both brightness and contrast to 50."
    
        clsProtocal.ResetPicMode
        clsProtocal.ChannelPreset

        If lvLastChk <= maxBrightnessSpec Then
            ShowError_Sys (30)
            GoTo FAIL
        End If
    End If

PASS:
    clsProtocal.ExitFacMode

    cmdMark = "PASS"
    Call saveALLcData

    CheckStep = CheckStep + "TEST ALL PASS"
    CheckStep.SelStart = Len(CheckStep)
    checkResult.ForeColor = &HC000&
    checkResult.Caption = "PASS"
    checkResult.BackColor = &HFF00&
    checkResult.ForeColor = &HC00000
    
    Label6.Caption = "PASS"
    
    Call subInitAfterRunning

    Exit Sub

FAIL:
    clsProtocal.ExitFacMode

    cmdMark = "FAIL"
    Call saveALLcData

    CheckStep.SelStart = Len(CheckStep)
    checkResult.BackColor = &HFF&
    checkResult.ForeColor = &H808080
    checkResult.Caption = "FAIL"
    checkResult.ForeColor = &H0&
    checkResult.ForeColor = &HFFFF&
    
    Label6.Caption = "FAIL"

    Call subInitAfterRunning

    Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub subInitBeforeRunning()
    countTime = 0
    lbTimer.Caption = "0s"
    Timer1.Enabled = True

    txtInput.Enabled = False
    'gstrBarCode = ""
    adjustGainAgainCool1Flag = 0
    adjustGainAgainNormalFlag = 0
    adjustGainAgainWarm1Flag = 0
End Sub

Private Sub subInitAfterRunning()
    Timer1.Enabled = False
    
    SaveLogInFile "[Time]Total: " & lbTimer.Caption & vbCrLf
    
    adjustGainAgainCool1Flag = 0
    adjustGainAgainNormalFlag = 0
    adjustGainAgainWarm1Flag = 0

    txtInput.Enabled = True
    txtInput.Text = ""
    txtInput.SetFocus
    
    If utdCommMode = modeNetwork Then
        isNetworkConnected = False
        tcpClient.Close
    End If
End Sub

Sub ShowError_Sys(t As Integer)
    Dim s As String
    
    s = "Unknown"

    Select Case t
        Case 1
            s = TXTGainCool1Wrong
        Case 2
            s = "ColorTemp_COOL_2 is Wrong, Please Check Again."
        Case 3
            s = TXTGainNormalWrong
        Case 4
            s = TXTGainWarm1Wrong
        Case 5
            s = "ColorTemp_WARM_2 is Wrong, Please Check Again."
        Case 6
            s = "LAB_SN:" + gstrBarCode + "(End)  Len:" + str$(gintBarCodeLen) + vbCrLf + "条形码长度不对，请确认！"
        Case 7
            s = "Can not Write DVI EDID."
        Case 8
            s = "Calibrate FAIL.(AUTO LEVEL)"
        Case 9
            s = "RS232 Connector Error"
        Case 10
            s = "Read DSUB EDID FAIL"
        Case 11
            s = TXTOffsetCool1Wrong
        Case 12
            s = "OFFSET_Color_COOL_2 is Wrong, Please Check Again."
        Case 13
            s = TXTOffsetNormalWrong
        Case 14
            s = TXTOffsetWarm1Wrong
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
            s = TXTLvTooLow
        Case 31
            s = ""
    End Select

    CheckStep.Text = CheckStep.Text + "Error Code:" + str$(t) + vbCrLf + s + vbCrLf
    CheckStep.SelStart = Len(CheckStep)
End Sub

Private Function AdjRGBGain(strColorTemp As String, adjustVal As Long, HighLowMode As Long) As Boolean
    Dim i, j, k As Integer

    Call clsProtocal.SelColorTemp(strColorTemp, setTVInputSource, setTVInputSourcePortNum)

    ' Set Offset first
    If adjustGainAgainCool1Flag = 0 Then
        Call setColorTemp(strColorTemp, presetData, 0)
        'DelayMS 200
        
        rRGB.cRR = presetData.nColorRR
        rRGB.cGG = presetData.nColorGG
        rRGB.cBB = presetData.nColorBB
        
        Call saveData(strColorTemp, 0)
    End If

    Call LoadData(strColorTemp, 0)
    If UCase(gstrChipSet) = "MST6M60" Then
        Call clsProtocal.SetRGBOffset(rRGB1.cRR * 8, rRGB1.cGG * 8, rRGB1.cBB * 8)
    Else
        Call clsProtocal.SetRGBOffset(rRGB1.cRR, rRGB1.cGG, rRGB1.cBB)
    End If
    
    Log_Info "========Adjust " & strColorTemp & "========"

    For j = 1 To 2
        Call setColorTemp(strColorTemp, presetData, HighLowMode)
        'DelayMS 200
        
        Log_Info "Init current colorTemp. RES:" + str$(RES)
        rRGB.cRR = presetData.nColorRR
        rRGB.cGG = presetData.nColorGG
        rRGB.cBB = presetData.nColorBB
        
        Label1 = CStr(presetData.xx)
        Label3 = CStr(presetData.yy)

        If UCase(gstrChipSet) = "MST6M60" Then
            Call clsProtocal.SetRGBGain(rRGB.cRR * 8, rRGB.cGG * 8, rRGB.cBB * 8)
        Else
            Call clsProtocal.SetRGBGain(rRGB.cRR, rRGB.cGG, rRGB.cBB)
        End If

        showData (1)

        resCodeForAdjustColorTemp = 0
        
        For k = 1 To 50
            If IsStop = True Then GoTo Cancel
            
            RES = checkColorTemp(rColor, strColorTemp)
            Log_Info "Check colorTemp. RES: " + CStr(RES)
            Log_Info "SPEC: x = " & CStr(presetData.xx) & " y = " & CStr(presetData.yy)
            Log_Info "Tol: x = " & CStr(presetData.xt) & " y =  " & CStr(presetData.yt)

            If RES = 3 Then
                Exit For
            Else
                If UCase(gstrBrand) = "CAN" Or _
                    UCase(gstrBrand) = "HAIER" Then
                    Call adjustColorTempForCIBN(rRGB)
                Else    ' Letv
                    If resCodeForAdjustColorTemp = 0 Then
                        Call adjustColorTemp(adjustMode3, rRGB, resCodeForAdjustColorTemp)
                    ElseIf resCodeForAdjustColorTemp = 1 Then
                        Call adjustColorTemp(adjustMode1, rRGB, resCodeForAdjustColorTemp)
                    ElseIf resCodeForAdjustColorTemp = 2 Then
                        Call adjustColorTemp(adjustMode2, rRGB, resCodeForAdjustColorTemp)
                    ElseIf resCodeForAdjustColorTemp = 3 Then
                        Call adjustColorTemp(adjustMode3, rRGB, resCodeForAdjustColorTemp)
                    ElseIf resCodeForAdjustColorTemp = 4 Then
                        Call adjustColorTemp(adjustMode4, rRGB, resCodeForAdjustColorTemp)
                    End If
                End If

                Log_Info "SET_RGB_GAN: R = " & CStr(rRGB.cRR) & _
                    ", G = " & CStr(rRGB.cGG) & ", B = " & CStr(rRGB.cBB) & _
                    ", resultcode = " & CStr(resCodeForAdjustColorTemp)

                If UCase(gstrChipSet) = "MST6M60" Then
                   Call clsProtocal.SetRGBGain(rRGB.cRR * 8, rRGB.cGG * 8, rRGB.cBB * 8)
                Else
                   Call clsProtocal.SetRGBGain(rRGB.cRR, rRGB.cGG, rRGB.cBB)
                End If

                showData (2)
            End If
        Next k
        
        If RES = 3 Then Exit For
        
    Next j

Cancel:
    If RES = 3 Then
        Call saveData(strColorTemp, HighLowMode)
        Log_Info "Save current data of " & strColorTemp & "."
        AdjRGBGain = True
    Else
        AdjRGBGain = False
    End If

End Function

Private Function AdjRGBOffset(strColorTemp As String, FixValue As Long, HighLowMode As Long) As Boolean
    Dim i, j, k As Integer

    Call clsProtocal.SelColorTemp(strColorTemp, setTVInputSource, setTVInputSourcePortNum)

    Log_Info "========Adjust " & strColorTemp & "========"
  
    For j = 1 To 2
        Call setColorTemp(strColorTemp, presetData, HighLowMode)
        'DelayMS 200
        Log_Info "Init current colorTemp. RES:" + str$(RES)
        rRGB.cRR = presetData.nColorRR
        rRGB.cGG = presetData.nColorGG
        rRGB.cBB = presetData.nColorBB
  
        'Label1 = Str$(presetData.xx)
        'Label3 = Str$(presetData.yy)

        Call clsProtocal.SetRGBOffset(rRGB.cRR, rRGB.cGG, rRGB.cBB)

        showData (3)

        For k = 1 To 50
            If IsStop = True Then GoTo Cancel
                
            RES = checkColorTemp(rColor, strColorTemp)
            Log_Info "Check colorTemp. RES:" + str$(RES)
    
            If RES = 3 Then
                Exit For
            Else
                Call adjustColorTempOffset(rRGB)
                    
                Log_Info "SET_RGB_OFFSET: R = " & CStr(rRGB.cRR) & _
                    ", G = " & CStr(rRGB.cGG) & ", B = " & CStr(rRGB.cBB)

                Call clsProtocal.SetRGBOffset(rRGB.cRR, rRGB.cGG, rRGB.cBB)
    
                showData (4)
            End If
        Next k

        If RES = 3 Then Exit For
    Next j

Cancel:
    If RES = 3 Then
        Call saveData(strColorTemp, HighLowMode)
        Log_Info "Save current data of " & strColorTemp & "."
        AdjRGBOffset = True
    Else
        AdjRGBOffset = False
    End If

End Function

Private Function checkColorAgain(strColorTemp As String, HighLowMode As Long) As Boolean
    Dim i, j, k As Integer

    Call clsProtocal.SelColorTemp(strColorTemp, setTVInputSource, setTVInputSourcePortNum)

    Log_Info "========Check " & strColorTemp & "========"
  
    For j = 1 To 2
        Call setColorTemp(strColorTemp, presetData, HighLowMode)
        'DelayMS 200
        Log_Info "Init current colorTemp. RES:" + str$(RES)

        Label1 = str$(presetData.xx)
        Label3 = str$(presetData.yy)

        showData (5)

        If IsStop = True Then GoTo Cancel

        RES = checkColorTemp(rColor, strColorTemp)
        Log_Info "Check colorTemp. RES:" + str$(RES)

        If RES = 3 Then Exit For
    Next j
  
Cancel:
    If RES = 3 Then
        checkColorAgain = True
    Else
        checkColorAgain = False
    End If

End Function


'step = lastChkShwDataStep: Check max brightness of TV with brightness 100 and contrast 100 in 100% white pattern.
Private Sub showData(step As Integer)
On Error Resume Next
    Dim xPos, yPos, vPos As Long

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
                RES = 0
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
 
    Log_Info "_x/y/Lv: " + CStr(rColor.xx) + " / " + CStr(rColor.yy) + " / " + CStr(rColor.lv)

    If Label6 <> "CHECK" Then Log_Info "_R/G/B: " + CStr(rRGB.cRR) + " / " + CStr(rRGB.cGG) + " / " + CStr(rRGB.cBB)

    Label_x = CStr(rColor.xx)
    Label_y = CStr(rColor.yy)
    Label_Lv = CStr(rColor.lv)
End Sub

Private Sub tbDisConnectastro_Click()
    If IsCa210ok Then
        ObjCa.RemoteMode = 0
    End If
End Sub

Private Sub Timer1_Timer()
    countTime = countTime + 1
    lbTimer.Caption = CStr(countTime) & "s"
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
    txtInput.Enabled = True
    
    Me.Caption = TXTTitle
    mTitle = Me.Caption
    subInitInterface

    gstrBrand = Split(gstrCurProjName, gstrDelimiterForProjName)(0)
    
    If UCase(gstrBrand) = "CAN" Then    'CANTV
        Set clsCANTVProtocal = New CANTVProtocal
        Set clsProtocal = clsCANTVProtocal
        PictureBrand.Picture = LoadPicture(App.Path & "\Resources\CANTV.bmp")
    ElseIf UCase(gstrBrand) = "HAIER" Then    'Haier
        Set clsHaierProtocal = New HaierProtocal
        Set clsProtocal = clsHaierProtocal
        PictureBrand.Picture = LoadPicture(App.Path & "\Resources\Haier.bmp")
    Else    'Letv
        If UCase(gstrChipSet) = "HX6310" Then
            Set clsLetvCurvedProtocal = New LetvCurvedProtocal
            Set clsProtocal = clsLetvCurvedProtocal
            PictureBrand.Picture = LoadPicture(App.Path & "\Resources\Letv.bmp")
        ElseIf UCase(gstrChipSet) = "MST6M60" Then
            Set clsLetvMST6M60 = New LetvMST6M60
            Set clsProtocal = clsLetvMST6M60
        Else
            Set clsLetvProtocal = New LetvProtocal
            Set clsProtocal = clsLetvProtocal
            PictureBrand.Picture = LoadPicture(App.Path & "\Resources\Letv.bmp")
        End If
    End If
    
    RES = initColorTemp(Calibrate, MinBrightness, gstrCurProjName, App.Path)
End Sub

Public Sub subInitInterface()
    Dim clsConfigData As ProjectConfig

    Set clsConfigData = New ProjectConfig
    clsConfigData.LoadConfigData
    
    setTVCurrentComBaud = clsConfigData.ComBaud
    setTVCurrentComID = clsConfigData.ComID
    glngI2cClockRate = clsConfigData.I2cClockRate
    setTVInputSource = clsConfigData.inputSource
    setTVInputSourcePortNum = CInt(Right(setTVInputSource, 1))
    setTVInputSource = Left(setTVInputSource, Len(setTVInputSource) - 1)
    delayTime = clsConfigData.DelayMS
    Ca210ChannelNO = clsConfigData.ChannelNum
    gintBarCodeLen = clsConfigData.BarCodeLen
    maxBrightnessSpec = clsConfigData.LvSpec
    gstrVPGModel = clsConfigData.VPGModel
    gstrVPGTiming = clsConfigData.VPGTiming
    gstrVPG100IRE = clsConfigData.VPG100IRE
    gstrVPG80IRE = clsConfigData.VPG80IRE
    gstrVPG20IRE = clsConfigData.VPG20IRE
    isAdjustCool2 = clsConfigData.EnableCool2
    isAdjustCool1 = clsConfigData.EnableCool1
    isAdjustNormal = clsConfigData.EnableNormal
    isAdjustWarm1 = clsConfigData.EnableWarm1
    isAdjustWarm2 = clsConfigData.EnableWarm2
    isCheckColorTemp = clsConfigData.EnableChkColor
    isAdjustOffset = clsConfigData.EnableAdjOffset
    gstrChipSet = clsConfigData.ChipSet
    
    utdCommMode = clsConfigData.CommMode
    If utdCommMode = modeUART Then
        subInitComPort
    ElseIf utdCommMode = modeNetwork Then
        subInitNetwork
    End If
    
    Set clsConfigData = Nothing

    txtInput.Text = ""
    lbModelName.Caption = Split(gstrCurProjName, gstrDelimiterForProjName)(1)
    
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
    
    InitVPGDevice
    DelayMS 200
    
    Call ChangeTiming(gstrVPGTiming)
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
        
        If txtInput.Enabled = True Then
            If txtInput.Text = "" Or Len(txtInput.Text) <> gintBarCodeLen Then
                MsgBox TXTBarcodeError & CStr(gintBarCodeLen), vbOKOnly, TXTBarcodeErrorTitle
                txtInput.Text = ""
                Exit Sub
            Else
                gstrBarCode = txtInput.Text
            End If

            SaveLogInFile "======================================================================="
            SaveLogInFile "        Auto-White Balance Adjusting Tool by Echom                     "
            SaveLogInFile "        Software Version: " & App.Major & "." & App.Minor & "." & App.Revision
            SaveLogInFile "        Barcode of TV: " & gstrBarCode
            SaveLogInFile "======================================================================="

            If utdCommMode = modeUART Then
                If MSComm1.PortOpen = False Then
                    MSComm1.PortOpen = True
                End If
                subMainProcesser
            ElseIf utdCommMode = modeNetwork Then
                isNetworkConnected = False
                Do
                    If tcpClient.State = sckClosed Then
                        Log_Info "TCP Connect"
                        tcpClient.Connect
                        txtInput.Enabled = False
                    End If
                    Call DelaySWithFlag(cmdReceiveWaitS * 2, isNetworkConnected)
                
                    If tcpClient.State = sckConnected Then
                        subMainProcesser
                        Exit Do
                    Else
                        If tcpClient.State <> sckClosed Then
                            tcpClient.Close
                        End If
                        i = i + 1
                    End If
                    Log_Info "Re-connect to TV."
                Loop While i <= 5
                txtInput.Enabled = True
            ElseIf utdCommMode = modeI2c Then
                Dim SetDeviceSts As Integer

                If DEVICE_USED = 0 Then
                    '=====================================
                    '  I2C tool initialization
                    '=====================================
                    SetDeviceSts = LptioSetDevice(DEVICE_FTDI)
    
                    '=====================================
                    '  Set I2C Clock Rate
                    '=====================================
                    Call I2cSetClockRate(glngI2cClockRate)
                    
                    DEVICE_USED = 1
                End If
                
                subMainProcesser
            End If
        End If
        
        If IsStop = True Then
            Exit Sub
        End If
    End If
    Exit Sub

ErrExit:
    txtInput.Text = ""
    MsgBox Err.Description, vbCritical, Err.Source
    'Invalid Port Number
    'If Err.Number = 8002 Then
    '    MsgBox Err.Description, vbCritical, Err.Source
    'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrExit

    If UCase(gstrBrand) = "CAN" Then
        If Not (clsCANTVProtocal Is Nothing) Then
            Set clsCANTVProtocal = Nothing
        End If
    ElseIf UCase(gstrBrand) = "HAIER" Then
        If Not (clsHaierProtocal Is Nothing) Then
            Set clsHaierProtocal = Nothing
        End If
    Else
        If UCase(gstrChipSet) = "HX6310" Then
            If Not (clsLetvCurvedProtocal Is Nothing) Then
                Set clsLetvCurvedProtocal = Nothing
            End If
        ElseIf UCase(gstrChipSet) = "MST6M60" Then
            If Not (clsLetvMST6M60 Is Nothing) Then
                Set clsLetvMST6M60 = Nothing
            End If
        Else
            If Not (clsLetvProtocal Is Nothing) Then
                Set clsLetvProtocal = Nothing
            End If
        End If
    End If
    
    If Not (clsProtocal Is Nothing) Then
        Set clsProtocal = Nothing
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


Private Sub saveData(strColorTemp As String, HL As Long)

    Select Case strColorTemp
        Case cstrColorTempCool1
            If HL Then
                cCOOL1.xx = rColor.xx
                cCOOL1.yy = rColor.yy
                cCOOL1.lv = rColor.lv
                cCOOL1.nColorRR = rRGB.cRR
                cCOOL1.nColorGG = rRGB.cGG
                cCOOL1.nColorBB = rRGB.cBB
            Else
                cFFCOOL1.xx = rColor.xx
                cFFCOOL1.yy = rColor.yy
                cFFCOOL1.lv = rColor.lv
                cFFCOOL1.nColorRR = rRGB.cRR
                cFFCOOL1.nColorGG = rRGB.cGG
                cFFCOOL1.nColorBB = rRGB.cBB
            End If

        Case cstrColorTempNormal
            If HL Then
                cNORMAL.xx = rColor.xx
                cNORMAL.yy = rColor.yy
                cNORMAL.lv = rColor.lv
                cNORMAL.nColorRR = rRGB.cRR
                cNORMAL.nColorGG = rRGB.cGG
                cNORMAL.nColorBB = rRGB.cBB
            Else
                cFFNORMAL.xx = rColor.xx
                cFFNORMAL.yy = rColor.yy
                cFFNORMAL.lv = rColor.lv
                cFFNORMAL.nColorRR = rRGB.cRR
                cFFNORMAL.nColorGG = rRGB.cGG
                cFFNORMAL.nColorBB = rRGB.cBB
            End If

        Case cstrColorTempWarm1
            If HL Then
                cWARM1.xx = rColor.xx
                cWARM1.yy = rColor.yy
                cWARM1.lv = rColor.lv
                cWARM1.nColorRR = rRGB.cRR
                cWARM1.nColorGG = rRGB.cGG
                cWARM1.nColorBB = rRGB.cBB
            Else
                cFFWARM1.xx = rColor.xx
                cFFWARM1.yy = rColor.yy
                cFFWARM1.lv = rColor.lv
                cFFWARM1.nColorRR = rRGB.cRR
                cFFWARM1.nColorGG = rRGB.cGG
                cFFWARM1.nColorBB = rRGB.cBB
            End If
    End Select
  
End Sub

Private Sub LoadData(strColorTemp As String, isGain As Boolean)
    Select Case strColorTemp
        Case cstrColorTempCool1
            If isGain Then
                rRGB1.cRR = cCOOL1.nColorRR
                rRGB1.cGG = cCOOL1.nColorGG
                rRGB1.cBB = cCOOL1.nColorBB
            Else
                rRGB1.cRR = cFFCOOL1.nColorRR
                rRGB1.cGG = cFFCOOL1.nColorGG
                rRGB1.cBB = cFFCOOL1.nColorBB
            End If
            
        Case cstrColorTempNormal
            If isGain Then
                rRGB1.cRR = cNORMAL.nColorRR
                rRGB1.cGG = cNORMAL.nColorGG
                rRGB1.cBB = cNORMAL.nColorBB
            Else
                rRGB1.cRR = cFFNORMAL.nColorRR
                rRGB1.cGG = cFFNORMAL.nColorGG
                rRGB1.cBB = cFFNORMAL.nColorBB
            End If
            
        Case cstrColorTempWarm1
            If isGain Then
                rRGB1.cRR = cWARM1.nColorRR
                rRGB1.cGG = cWARM1.nColorGG
                rRGB1.cBB = cWARM1.nColorBB
            Else
                rRGB1.cRR = cFFWARM1.nColorRR
                rRGB1.cGG = cFFWARM1.nColorGG
                rRGB1.cBB = cFFWARM1.nColorBB
            End If
    End Select
End Sub

Private Sub saveALLcData()
    If gstrBarCode = "" Then
        Exit Sub
    Else
        sqlstring = "select * from [" & gstrCurProjName & "]"
        Executesql (sqlstring)
        rs.AddNew

        rs.Fields(0) = gstrCurProjName
        rs.Fields(1) = gstrBarCode

        rs.Fields(2) = cCOOL1.xx
        rs.Fields(3) = cCOOL1.yy
        rs.Fields(4) = cCOOL1.nColorRR
        rs.Fields(5) = cCOOL1.nColorGG
        rs.Fields(6) = cCOOL1.nColorBB
        rs.Fields(7) = cNORMAL.xx
        rs.Fields(8) = cNORMAL.yy
        rs.Fields(9) = cNORMAL.nColorRR
        rs.Fields(10) = cNORMAL.nColorGG
        rs.Fields(11) = cNORMAL.nColorBB
        rs.Fields(12) = cWARM1.xx
        rs.Fields(13) = cWARM1.yy
        rs.Fields(14) = cWARM1.nColorRR
        rs.Fields(15) = cWARM1.nColorGG
        rs.Fields(16) = cWARM1.nColorBB
        
        rs.Fields(17) = cFFCOOL1.nColorRR
        rs.Fields(18) = cFFCOOL1.nColorGG
        rs.Fields(19) = cFFCOOL1.nColorBB
        rs.Fields(20) = cFFNORMAL.nColorRR
        rs.Fields(21) = cFFNORMAL.nColorGG
        rs.Fields(22) = cFFNORMAL.nColorBB
        rs.Fields(23) = cFFWARM1.nColorRR
        rs.Fields(24) = cFFWARM1.nColorGG
        rs.Fields(25) = cFFWARM1.nColorBB

        rs.Fields(26) = lvLastChk
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

Private Sub InitVPGDevice()
    Select Case gstrVPGModel
        Case "2401"
            Set ivpg = New VPGCtrl.VPGCtrl_24xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG2401)
        Case "2402"
            Set ivpg = New VPGCtrl.VPGCtrl_24xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG2402)
        Case "2333_B"
            Set ivpg = New VPGCtrl.VPGCtrl_24xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG2333_B)
        Case "23293_B"
            Set ivpg = New VPGCtrl.VPGCtrl_24xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG23293_B)
        Case "23294"
            Set ivpg = New VPGCtrl.VPGCtrl_24xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG23294)
        Case "22293"
            Set ivpg = New VPGCtrl.VPGCtrl_22xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG22293)
        Case "22293_A"
            Set ivpg = New VPGCtrl.VPGCtrl_22xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG22293_A)
        Case "22293_B"
            Set ivpg = New VPGCtrl.VPGCtrl_22xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG22293_B)
        Case "2233"
            Set ivpg = New VPGCtrl.VPGCtrl_22xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG2233)
        Case "2233_A"
            Set ivpg = New VPGCtrl.VPGCtrl_22xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG2233_A)
        Case "2233_B"
            Set ivpg = New VPGCtrl.VPGCtrl_22xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG2233_B)
        Case "2234"
            Set ivpg = New VPGCtrl.VPGCtrl_22xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG2234)
        Case "22294"
            Set ivpg = New VPGCtrl.VPGCtrl_22xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG22294)
        Case "22294_A"
            Set ivpg = New VPGCtrl.VPGCtrl_22xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG22294_A)
    End Select

End Sub

Private Sub Obj_OnChangedConnectState(ByVal bIsConnected As Boolean)
    If bIsConnected = False Then
        Me.Caption = mTitle & " [Chroma " & gstrVPGModel & " Disconnected]"
    Else
        Me.Caption = mTitle
    End If
End Sub

Private Sub ChangeTiming(Tim As String)
    Dim bNo(1) As Byte
    
    bNo(0) = (CInt(Tim) And &HFF00) \ 256
    bNo(1) = CInt(Tim) And &HFF

    ivpg.ExecuteCmd VPG_CMD_CM_DOWNLOAD, VPG_SCMD_SCM_CTL_RUNTIM, bNo, False
End Sub

Private Sub ChangePattern(Ptn As String)
    Dim bNo(1) As Byte
    
    bNo(0) = (CInt(Ptn) And &HFF00) \ 256
    bNo(1) = CInt(Ptn) And &HFF

    ivpg.RunKey (VPG_KEY_CKEY_OUT)
    ivpg.ExecuteCmd VPG_CMD_CM_DOWNLOAD, VPG_SCMD_SCM_CTL_RUNPTN, bNo, False
End Sub
