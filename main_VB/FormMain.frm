VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FormMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto White Balance Tool"
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
   Icon            =   "FormMain.frx":0000
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
      Picture         =   "FormMain.frx":1DF72
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
      Picture         =   "FormMain.frx":24226
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
      Text            =   "FormMain.frx":43320
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
   Begin VB.Label lbAdjustStandard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STANDARD"
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
Attribute VB_Name = "FormMain"
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
Dim resCodeForAdjustColorTemp As Long
Dim cmdMark As String
Dim clsProtocal As Protocal
Dim clsCANTVProtocal As CANTVProtocal
Dim clsLetvProtocal As LetvProtocal
Dim clsLetvCurvedProtocal As LetvCurvedProtocal
Dim clsLetvMST6M60 As LetvMST6M60
Dim clsHaierProtocal As HaierProtocal

Dim ivpg As IVPGCtrl

Private rRGB As REALRGB
Private rRGB1 As REALRGB

Private mAdjGainAgainCool1 As Integer
Private mAdjGainAgainStandard As Integer
Private mAdjGainAgainWarm1 As Integer
Private mCntTime As Long
Private mTitle As String
Private mBrand As String
Private mBarCode As String

Private WithEvents Obj As VPGCtrl.VPGCtrl
Attribute Obj.VB_VarHelpID = -1

Private Sub SubRun()
    On Error GoTo ErrExit
    subInitBeforeRunning

    If gblnStop = True Then
        Exit Sub
    End If

    If gblnCaConnected = False Then
        MsgBox TXTCaDisconnectHint, vbOKOnly + vbInformation, "warning"
        subInitAfterRunning
        
        Exit Sub
    End If

    checkResult.BackColor = &H80FFFF
    gblnStop = False
    checkResult.Caption = TXTRun
    checkResult.ForeColor = &HC0&
    CheckStep = ""

    lbAdjustCOOL_1.BackColor = &H8000000F
    lbAdjustCOOL_2.BackColor = &H8000000F
    lbAdjustStandard.BackColor = &H8000000F
    lbAdjustWARM_1.BackColor = &H8000000F
    lbAdjustWARM_2.BackColor = &H8000000F

    Picture1.Cls
    lbColorTempWrong.Visible = False

    Set ObjMemory = ObjCa.Memory
    ObjMemory.ChannelNO = glngCaChannel

    SubLogInfo "Start adjusting color temperature"
    Call ChangePattern(gstrVPG80IRE)

    clsProtocal.EnterFacMode
    Call clsProtocal.SwitchInputSource(gstrTvInputSrc, gintTvInputSrcPort)
    Call clsProtocal.ResetPicMode
    Call clsProtocal.SetBacklight(100)
    SubLogInfo "Set backlight to 100"

    Label6.Caption = "WHITE"

ADJUST_GAIN_AGAIN_COOL1:
    If gblnEnableCool1 Then
        lbAdjustCOOL_1.BackColor = &H80FFFF
        Result = FuncAdjRGBGain(COLORTEMP_COOL1, ADJMODE_3)
  
        If Result = False Then
            ShowError_Sys (1)
            GoTo FAIL
        Else
            Call clsProtocal.SaveWBDataToAllSrc(gstrTvInputSrc, gintTvInputSrcPort)
        End If

        SubSaveLogInFile "[Time]White Cool1: " & lbTimer.Caption
        lbAdjustCOOL_1.BackColor = &HC0FFC0
        
        If mAdjGainAgainCool1 > 0 Then
            GoTo CHECK_COOL1
        End If
    End If

ADJUST_GAIN_AGAIN_NORMAL:
    If gblnEnableStandard Then
        lbAdjustStandard.BackColor = &H80FFFF
        Result = FuncAdjRGBGain(COLORTEMP_STANDARD, ADJMODE_3)

        If Result = False Then
            ShowError_Sys (3)
            GoTo FAIL
        Else
            Call clsProtocal.SaveWBDataToAllSrc(gstrTvInputSrc, gintTvInputSrcPort)
        End If

        SubSaveLogInFile "[Time]White Normal: " & lbTimer.Caption
        lbAdjustStandard.BackColor = &HC0FFC0
        
        If mAdjGainAgainStandard > 0 Then
            GoTo CHECK_NORMAL
        End If
    End If

ADJUST_GAIN_AGAIN_WARM1:
    If gblnEnableWarm1 Then
        lbAdjustWARM_1.BackColor = &H80FFFF
        Result = FuncAdjRGBGain(COLORTEMP_WARM1, ADJMODE_3)

        If Result = False Then
            ShowError_Sys (4)
            GoTo FAIL
        Else
            Call clsProtocal.SaveWBDataToAllSrc(gstrTvInputSrc, gintTvInputSrcPort)
        End If

        SubSaveLogInFile "[Time]White Warm1: " & lbTimer.Caption
        lbAdjustWARM_1.BackColor = &HC0FFC0
        
        If mAdjGainAgainWarm1 > 0 Then
            GoTo CHECK_WARM1
        End If
    End If

    If gblnAdjOffset Then
        Label6.Caption = TXTGrey

        Call ChangePattern(gstrVPG20IRE)

        If gblnEnableCool1 Then
            lbAdjustCOOL_1.BackColor = &H80FFFF
            Result = FuncAdjRGBOffset(COLORTEMP_COOL1)
                
            If Result = False Then
                ShowError_Sys (11)
                GoTo FAIL
            Else
                Call clsProtocal.SaveWBDataToAllSrc(gstrTvInputSrc, gintTvInputSrcPort)
            End If
            
            SubSaveLogInFile "[Time]Grey Cool1: " & lbTimer.Caption
            lbAdjustCOOL_1.BackColor = &HC0FFC0
        End If
   
        If gblnEnableStandard Then
            lbAdjustStandard.BackColor = &H80FFFF
            Result = FuncAdjRGBOffset(COLORTEMP_STANDARD)

            If Result = False Then
                ShowError_Sys (13)
                GoTo FAIL
            Else
                Call clsProtocal.SaveWBDataToAllSrc(gstrTvInputSrc, gintTvInputSrcPort)
            End If

            SubSaveLogInFile "[Time]Grey Normal: " & lbTimer.Caption
            lbAdjustStandard.BackColor = &HC0FFC0
        End If
   
        If gblnEnableWarm1 Then
            lbAdjustWARM_1.BackColor = &H80FFFF
            Result = FuncAdjRGBOffset(COLORTEMP_WARM1)
                
            If Result = False Then
                ShowError_Sys (14)
                GoTo FAIL
            Else
                Call clsProtocal.SaveWBDataToAllSrc(gstrTvInputSrc, gintTvInputSrcPort)
            End If

            SubSaveLogInFile "[Time]Grey Warm1: " & lbTimer.Caption
            lbAdjustWARM_1.BackColor = &HC0FFC0
        End If
    End If

    If gblnChkColorTemp Then
        If gblnAdjOffset Then
            Call ChangePattern(gstrVPG80IRE)
        End If

CHECK_COOL1:
        If gblnEnableCool1 Then
            Label6.Caption = TXTChk
            lbAdjustCOOL_1.BackColor = &H80FFFF
            Result = checkColorAgain(COLORTEMP_COOL1)

            If Result = False Then
                ShowError_Sys (1)

                If mAdjGainAgainCool1 > 0 Then
                    GoTo FAIL
                End If
                
                mAdjGainAgainCool1 = mAdjGainAgainCool1 + 1
                
                GoTo ADJUST_GAIN_AGAIN_COOL1
            End If
      
            lbAdjustCOOL_1.BackColor = &HC0FFC0
        End If

CHECK_NORMAL:
        If gblnEnableStandard Then
            Label6.Caption = TXTChk
            lbAdjustStandard.BackColor = &H80FFFF
            Result = checkColorAgain(COLORTEMP_STANDARD)

            If Result = False Then
                ShowError_Sys (3)

                If mAdjGainAgainStandard > 0 Then
                    GoTo FAIL
                End If
    
                mAdjGainAgainStandard = mAdjGainAgainStandard + 1

                GoTo ADJUST_GAIN_AGAIN_NORMAL
            End If
    
            lbAdjustStandard.BackColor = &HC0FFC0
        End If

CHECK_WARM1:
        If gblnEnableWarm1 Then
            Label6.Caption = TXTChk
            lbAdjustWARM_1.BackColor = &H80FFFF
            Result = checkColorAgain(COLORTEMP_WARM1)

            If Result = False Then
                ShowError_Sys (4)
                
                If mAdjGainAgainWarm1 > 0 Then
                    GoTo FAIL
                End If
    
                mAdjGainAgainWarm1 = mAdjGainAgainWarm1 + 1
                
                GoTo ADJUST_GAIN_AGAIN_WARM1
            End If

            lbAdjustWARM_1.BackColor = &HC0FFC0
        End If
    End If
    
    If gstrChipSet = "T111" Then
        Call clsProtocal.SelColorTemp(COLORTEMP_STANDARD, gstrTvInputSrc, gintTvInputSrcPort)
        SubLogInfo "Set color temp to cool1"
        
        ObjCa.Measure
        lvLastChk = CLng(ObjProbe.lv)
        SubLogInfo "lv = " + CStr(lvLastChk)
        SubShowData (LASTSTEP)
        
        If lvLastChk <= glngBlSpecVal Then
            ShowError_Sys (30)
            GoTo FAIL
        End If
    Else
        'Last check:
        'Cool, 100% white pattern, brightness = 100, contrast = 100
        'Check Lv and save x, y, lv
        Call ChangePattern(gstrVPG100IRE)

        Call clsProtocal.SetBrightness(100)
        SubLogInfo "Set brightness to 100"

        Call clsProtocal.SetContrast(100)
        SubLogInfo "Set contrast to 100"

        Call clsProtocal.SelColorTemp(COLORTEMP_COOL1, gstrTvInputSrc, gintTvInputSrcPort)
        SubLogInfo "Set color temp to cool1"

        ObjCa.Measure
        lvLastChk = CLng(ObjProbe.lv)
        SubLogInfo "lv = " + CStr(lvLastChk)
        SubShowData (LASTSTEP)

        Call clsProtocal.SetBrightness(50)
        Call clsProtocal.SetContrast(50)
        SubLogInfo "Set both brightness and contrast to 50."
    
        clsProtocal.ResetPicMode
        clsProtocal.ChannelPreset

        If lvLastChk <= glngBlSpecVal Then
            ShowError_Sys (30)
            GoTo FAIL
        End If
    End If

PASS:
    clsProtocal.ExitFacMode

    cmdMark = TXTPass
    Call saveALLcData

    CheckStep = CheckStep + "TEST ALL PASS"
    CheckStep.SelStart = Len(CheckStep)
    checkResult.ForeColor = &HC000&
    checkResult.Caption = TXTPass
    checkResult.BackColor = &HFF00&
    checkResult.ForeColor = &HC00000
    
    Label6.Caption = TXTPass
    
    Call subInitAfterRunning

    Exit Sub

FAIL:
    clsProtocal.ExitFacMode

    cmdMark = TXTFail
    Call saveALLcData

    CheckStep.SelStart = Len(CheckStep)
    checkResult.BackColor = &HFF&
    checkResult.ForeColor = &H808080
    checkResult.Caption = TXTFail
    checkResult.ForeColor = &H0&
    checkResult.ForeColor = &HFFFF&
    
    Label6.Caption = TXTFail

    Call subInitAfterRunning

    Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub subInitBeforeRunning()
    mCntTime = 0
    lbTimer.Caption = "0s"
    Timer1.Enabled = True

    txtInput.Enabled = False
    'mBarCode = ""
    mAdjGainAgainCool1 = 0
    mAdjGainAgainStandard = 0
    mAdjGainAgainWarm1 = 0
End Sub

Private Sub subInitAfterRunning()
    Timer1.Enabled = False
    
    SubSaveLogInFile "[Time]Total: " & lbTimer.Caption & vbCrLf
    
    mAdjGainAgainCool1 = 0
    mAdjGainAgainStandard = 0
    mAdjGainAgainWarm1 = 0

    txtInput.Enabled = True
    txtInput.Text = ""
    txtInput.SetFocus
    
    If gutdCommMode = modeNetwork Then
        gblnNetConnected = False
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
            s = TXTGainCool2Wrong
        Case 3
            s = TXTGainNormalWrong
        Case 4
            s = TXTGainWarm1Wrong
        Case 5
            s = TXTGainWarm2Wrong
        Case 6
            s = "LAB_SN:" + mBarCode + "(End)  Len:" + str$(gintBarCodeLen) + vbCrLf + TXTSNLenWrong
        Case 7
            s = TXTDVIWrong
        Case 8
            s = TXTCalFail
        Case 9
            s = TXTRS232Er
        Case 10
            s = TXTDSUBFail
        Case 11
            s = TXTOffsetCool1Wrong
        Case 12
            s = TXTOffsetCool2Wrong
        Case 13
            s = TXTOffsetNormalWrong
        Case 14
            s = TXTOffsetWarm1Wrong
        Case 15
            s = TXTOffsetWarm2Wrong
        Case 16
            s = TXTHDMI2ChkWrong
        Case 17
            s = TXTHDMI2EDIDWrong
        Case 18
            s = TXTMinBriTooHigh
        Case 19
            s = TXTFWVerWrong
        Case 20
            s = TXTOSDSNWriteWrong
        Case 21
            s = TXTMaxBriTooHigh
        Case 22
            s = TXTCT5000Wrong
        Case 23
            s = TXTCT3000Wrong
        Case 24
            s = TXTLSDataWrong
        Case 25
            s = TXTLvTooLow
        Case 26
            s = ""
    End Select

    CheckStep.Text = CheckStep.Text + TXTErrCode + str$(t) + vbCrLf + s + vbCrLf
    CheckStep.SelStart = Len(CheckStep)
End Sub

Private Function FuncAdjRGBGain(strColorTemp As String, adjustVal As Long) As Boolean
    Dim i, j As Integer

    Call clsProtocal.SelColorTemp(strColorTemp, gstrTvInputSrc, gintTvInputSrcPort)

    ' Set Offset first
    If mAdjGainAgainCool1 = 0 Then
        Call ColorTSetSpec(strColorTemp, presetData, 0)
        'SubDelayMs 200
        
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
    
    SubLogInfo "========Adjust " & strColorTemp & "========"

    For i = 1 To 2
        Call ColorTSetSpec(strColorTemp, presetData, ADJMODE_GAIN)
        'SubDelayMs 200
        
        SubLogInfo "Init current colorTemp. RES:" + str$(RES)
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

        SubShowData (1)

        resCodeForAdjustColorTemp = 0
        
        For j = 1 To 50
            If gblnStop = True Then GoTo Cancel
            
            RES = ColorTChk(rColor, strColorTemp)
            SubLogInfo "Check colorTemp. RES: " + CStr(RES)
            SubLogInfo "SPEC: x = " & CStr(presetData.xx) & " y = " & CStr(presetData.yy)
            SubLogInfo "Tol: x = " & CStr(presetData.xt) & " y =  " & CStr(presetData.yt)

            If RES = 3 Then
                Exit For
            Else
                If UCase(mBrand) = "CAN" Or _
                    UCase(mBrand) = "HAIER" Then
                    Call ColorTAdjRGBGain(rRGB)
                Else    ' Letv
                    If resCodeForAdjustColorTemp = 0 Then
                        Call ColorTAdjRGBGainLetv(ADJMODE_3, rRGB, resCodeForAdjustColorTemp)
                    ElseIf resCodeForAdjustColorTemp = 1 Then
                        Call ColorTAdjRGBGainLetv(ADJMODE_1, rRGB, resCodeForAdjustColorTemp)
                    ElseIf resCodeForAdjustColorTemp = 2 Then
                        Call ColorTAdjRGBGainLetv(ADJMODE_2, rRGB, resCodeForAdjustColorTemp)
                    ElseIf resCodeForAdjustColorTemp = 3 Then
                        Call ColorTAdjRGBGainLetv(ADJMODE_3, rRGB, resCodeForAdjustColorTemp)
                    ElseIf resCodeForAdjustColorTemp = 4 Then
                        Call ColorTAdjRGBGainLetv(ADJMODE_4, rRGB, resCodeForAdjustColorTemp)
                    End If
                End If

                SubLogInfo "SET_RGB_GAN: R = " & CStr(rRGB.cRR) & _
                    ", G = " & CStr(rRGB.cGG) & ", B = " & CStr(rRGB.cBB) & _
                    ", resultcode = " & CStr(resCodeForAdjustColorTemp)

                If UCase(gstrChipSet) = "MST6M60" Then
                   Call clsProtocal.SetRGBGain(rRGB.cRR * 8, rRGB.cGG * 8, rRGB.cBB * 8)
                Else
                   Call clsProtocal.SetRGBGain(rRGB.cRR, rRGB.cGG, rRGB.cBB)
                End If

                SubShowData (2)
            End If
        Next j
        
        If RES = 3 Then Exit For
        
    Next i

Cancel:
    If RES = 3 Then
        Call saveData(strColorTemp, ADJMODE_GAIN)
        SubLogInfo "Save current data of " & strColorTemp & "."
        FuncAdjRGBGain = True
    Else
        FuncAdjRGBGain = False
    End If

End Function

Private Function FuncAdjRGBOffset(strColorTemp As String) As Boolean
    Dim i, j As Integer

    Call clsProtocal.SelColorTemp(strColorTemp, gstrTvInputSrc, gintTvInputSrcPort)

    SubLogInfo "========Adjust " & strColorTemp & "========"
  
    For i = 1 To 2
        Call ColorTSetSpec(strColorTemp, presetData, ADJMODE_OFFSET)
        'SubDelayMs 200
        SubLogInfo "Init current colorTemp. RES:" + str$(RES)
        rRGB.cRR = presetData.nColorRR
        rRGB.cGG = presetData.nColorGG
        rRGB.cBB = presetData.nColorBB
  
        'Label1 = Str$(presetData.xx)
        'Label3 = Str$(presetData.yy)

        Call clsProtocal.SetRGBOffset(rRGB.cRR, rRGB.cGG, rRGB.cBB)

        SubShowData (3)

        For j = 1 To 50
            If gblnStop = True Then GoTo Cancel
                
            RES = ColorTChk(rColor, strColorTemp)
            SubLogInfo "Check colorTemp. RES:" + str$(RES)
    
            If RES = 3 Then
                Exit For
            Else
                Call ColorTAdjRGBOffset(rRGB)
                    
                SubLogInfo "SET_RGB_OFFSET: R = " & CStr(rRGB.cRR) & _
                    ", G = " & CStr(rRGB.cGG) & ", B = " & CStr(rRGB.cBB)

                Call clsProtocal.SetRGBOffset(rRGB.cRR, rRGB.cGG, rRGB.cBB)
    
                SubShowData (4)
            End If
        Next j

        If RES = 3 Then Exit For
    Next i

Cancel:
    If RES = 3 Then
        Call saveData(strColorTemp, ADJMODE_OFFSET)
        SubLogInfo "Save current data of " & strColorTemp & "."
        FuncAdjRGBOffset = True
    Else
        FuncAdjRGBOffset = False
    End If

End Function

Private Function checkColorAgain(strColorTemp As String) As Boolean
    Dim i As Integer

    Call clsProtocal.SelColorTemp(strColorTemp, gstrTvInputSrc, gintTvInputSrcPort)

    SubLogInfo "========Check " & strColorTemp & "========"
  
    For i = 1 To 2
        Call ColorTSetSpec(strColorTemp, presetData, ADJMODE_GAIN)
        'SubDelayMs 200
        SubLogInfo "Init current colorTemp. RES:" + str$(RES)

        Label1 = str$(presetData.xx)
        Label3 = str$(presetData.yy)

        SubShowData (5)

        If gblnStop = True Then GoTo Cancel

        RES = ColorTChk(rColor, strColorTemp)
        SubLogInfo "Check colorTemp. RES:" + str$(RES)

        If RES = 3 Then Exit For
    Next i
  
Cancel:
    If RES = 3 Then
        checkColorAgain = True
    Else
        checkColorAgain = False
    End If

End Function


'step = LASTSTEP: Check max brightness of TV with brightness 100 and contrast 100 in 100% white pattern.
Private Sub SubShowData(step As Integer)
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
    If presetData.xt = 0 Then
        presetData.xt = 30
    End If
    If presetData.yt = 0 Then
        presetData.yt = 30
    End If
    xPos = 1515 + (rColor.xx - presetData.xx) * 365 / presetData.xt
    yPos = 1275 - (rColor.yy - presetData.yy) * 385 / presetData.yt

    If step = LASTSTEP Then
        vPos = 1660 - (rColor.lv - glngBlSpecVal) * 385 / 50
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

    If step <> LASTSTEP Then
        If Abs(rColor.xx - presetData.xx) <= presetData.xt And Abs(rColor.yy - presetData.yy) <= presetData.yt Then
            lbColorTempWrong.Visible = False
            Picture1.Circle (xPos, yPos), 23, &H30FF30
        Else
            lbColorTempWrong.Visible = True
            Picture1.Circle (xPos, yPos), 23, &HFF&

            If rColor.xx < 5 Then
                gblnStop = True
                ObjCa.RemoteMode = 2
                MsgBox (TXTChkCA210)
                RES = 0
            End If
        End If
    End If

    'In lv axis, 3060 is the distance from left edge of white rectangle to the left of Picture1.
    'In lv axis, 3390 is the distance from right edge of white rectangle to the left of Picture1.
    If step = LASTSTEP Then
        If rColor.lv > glngBlSpecVal Then
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
 
    SubLogInfo "_x/y/Lv: " + CStr(rColor.xx) + " / " + CStr(rColor.yy) + " / " + CStr(rColor.lv)

    If Label6 <> TXTChk Then SubLogInfo "_R/G/B: " + CStr(rRGB.cRR) + " / " + CStr(rRGB.cGG) + " / " + CStr(rRGB.cBB)

    Label_x = CStr(rColor.xx)
    Label_y = CStr(rColor.yy)
    Label_Lv = CStr(rColor.lv)
End Sub

Private Sub tbDisConnectastro_Click()
    If gblnCaConnected Then
        ObjCa.RemoteMode = 0
    End If
End Sub

Private Sub Timer1_Timer()
    mCntTime = mCntTime + 1
    lbTimer.Caption = CStr(mCntTime) & "s"
End Sub

Private Sub vbSetSPEC_Click()
    FormSettings.Show
End Sub

Private Sub vbAbout_Click()
    FormAbout.Show
End Sub

Private Sub vbConCA310_Click()
    If gblnCaConnected = True Then
        ObjCa.RemoteMode = 1
        Exit Sub
    Else
        SubConnectCa
    End If
End Sub


Private Sub Form_Load()
    vbFunc.Caption = TXTFun
    vbConCA310.Caption = TXTConnectCA
    tbDisConnectastro.Caption = TXTDisConnectCA
    vbSet.Caption = TXTSet
    vbSetSPEC.Caption = TXTSetSpec
    vbDescription.Caption = TXTDiscription
    vbAbout.Caption = TXTAbout
    lbAdjustCOOL_1.Caption = TXTCOOL1
    lbAdjustCOOL_2.Caption = TXTCOOL2
    lbAdjustStandard.Caption = TXTSTD
    lbAdjustWARM_1.Caption = TXTWARM1
    lbAdjustWARM_2.Caption = TXTWARM2
    Label6.Caption = TXTINITIAL
    Label7.Caption = "SPEC"
    checkResult.Caption = TXTChkResult
    gblnStop = False
    txtInput.Enabled = True
    
    Me.Caption = TXTTitle & " V" & App.Major & "." & App.Minor & "." & App.Revision
    mTitle = Me.Caption
    subInitInterface

    mBrand = Split(gstrCurProjName, gstrDelimiterForProjName)(0)
    
    If UCase(mBrand) = "CAN" Then    'CANTV
        Set clsCANTVProtocal = New CANTVProtocal
        Set clsProtocal = clsCANTVProtocal
        PictureBrand.Picture = LoadPicture(App.Path & "\Resources\CANTV.bmp")
    ElseIf UCase(mBrand) = "HAIER" Then    'Haier
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
    
    RES = ColorTInit(rConfigData)
End Sub

Public Sub subInitInterface()
    
    LoadConfigData
    LoadConfigData1
    
    gintCurComBaud = ComBaud
    gintCurComId = ComID
    glngI2cClockRate = I2cClockRate
    gstrTvInputSrc = inputSource
    gintTvInputSrcPort = CInt(Right(gstrTvInputSrc, 1))
    gstrTvInputSrc = Left(gstrTvInputSrc, Len(gstrTvInputSrc) - 1)
    glngDelayTime = DelayMS
    glngCaChannel = ChannelNum
    gintBarCodeLen = BarCodeLen
    glngBlSpecVal = LvSpec
    gstrVPGModel = VPGModel
    gstrVPGTiming = VPGTiming
    gstrVPG100IRE = VPG100IRE
    gstrVPG80IRE = VPG80IRE
    gstrVPG20IRE = VPG20IRE
    gblnEnableCool2 = EnableCool2
    gblnEnableCool1 = EnableCool1
    gblnEnableStandard = EnableNormal
    gblnEnableWarm1 = EnableWarm1
    gblnEnableWarm2 = EnableWarm2
    gblnChkColorTemp = EnableChkColor
    gblnAdjOffset = EnableAdjOffset
    gstrChipSet = ChipSet
    
    gutdCommMode = CommMode
    If gutdCommMode = modeUART Then
        subInitComPort
    ElseIf gutdCommMode = modeNetwork Then
        subInitNetwork
    End If
    

    txtInput.Text = ""
    lbModelName.Caption = Split(gstrCurProjName, gstrDelimiterForProjName)(1)
    
    If gblnEnableCool1 = True Then lbAdjustCOOL_1.ForeColor = &H80000008
    If gblnEnableCool2 = True Then lbAdjustCOOL_2.ForeColor = &H80000008
    If gblnEnableStandard = True Then lbAdjustStandard.ForeColor = &H80000008
    If gblnEnableWarm1 = True Then lbAdjustWARM_1.ForeColor = &H80000008
    If gblnEnableWarm2 = True Then lbAdjustWARM_2.ForeColor = &H80000008

    If gblnEnableCool1 = False Then lbAdjustCOOL_1.ForeColor = &HC0C0C0
    If gblnEnableCool2 = False Then lbAdjustCOOL_2.ForeColor = &HC0C0C0
    If gblnEnableStandard = False Then lbAdjustStandard.ForeColor = &HC0C0C0
    If gblnEnableWarm1 = False Then lbAdjustWARM_1.ForeColor = &HC0C0C0
    If gblnEnableWarm2 = False Then lbAdjustWARM_2.ForeColor = &HC0C0C0
    
    InitVPGDevice
    SubDelayMs 200
    
    Call ChangeTiming(gstrVPGTiming)
End Sub

Private Sub subInitComPort()
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
    
    MSComm1.CommPort = gintCurComId
    MSComm1.Settings = gintCurComBaud & ",N,8,1"
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
    gblnNetConnected = False
    With tcpClient
        .Protocol = sckTCPProtocol
        ' IMPORTANT: be sure to change the RemoteHost
        ' value to the name of your computer.
        .RemoteHost = REMOTE_HOST
        .RemotePort = REMOTE_PORT
    End With
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrExit
    Dim i As Integer
    
    i = 0

    If KeyAscii = 13 Then
        gblnStop = False
        
        If txtInput.Enabled = True Then
            If txtInput.Text = "" Or Len(txtInput.Text) <> gintBarCodeLen Then
                MsgBox TXTBarcodeError & CStr(gintBarCodeLen), vbOKOnly, TXTBarcodeErrorTitle
                txtInput.Text = ""
                Exit Sub
            Else
                mBarCode = txtInput.Text
            End If

            SubSaveLogInFile "======================================================================="
            SubSaveLogInFile "        Auto-White Balance Adjusting Tool by Echom                     "
            SubSaveLogInFile "        Software Version: " & App.Major & "." & App.Minor & "." & App.Revision
            SubSaveLogInFile "        Barcode of TV: " & mBarCode
            SubSaveLogInFile "======================================================================="

            If gutdCommMode = modeUART Then
                If MSComm1.PortOpen = False Then
                    MSComm1.PortOpen = True
                End If
                SubRun
            ElseIf gutdCommMode = modeNetwork Then
                gblnNetConnected = False
                Do
                    If tcpClient.State = sckClosed Then
                        SubLogInfo "TCP Connect"
                        tcpClient.Connect
                        txtInput.Enabled = False
                    End If
                    Call SubDelayWithFlag(10, gblnNetConnected)
                
                    If tcpClient.State = sckConnected Then
                        SubRun
                        Exit Do
                    Else
                        If tcpClient.State <> sckClosed Then
                            tcpClient.Close
                        End If
                        i = i + 1
                    End If
                    SubLogInfo "Re-connect to TV."
                Loop While i <= 5
                txtInput.Enabled = True
            ElseIf gutdCommMode = modeI2c Then
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
                
                SubRun
            End If
        End If
        
        If gblnStop = True Then
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
    Dim strColorTemp As String
    Dim isGain As Boolean
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    success = xmlDoc.Load(gstrXmlPath)

On Error GoTo ErrExit

    If UCase(mBrand) = "CAN" Then
        If Not (clsCANTVProtocal Is Nothing) Then
            Set clsCANTVProtocal = Nothing
        End If
    ElseIf UCase(mBrand) = "HAIER" Then
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

    gblnStop = True
    If (gblnCaConnected = True) Then
        ObjCa.RemoteMode = 0
    End If
  
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
  
    Call ColorTDeInit(rConfigData)
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        Select Case strColorTemp
        Case COLORTEMP_COOL1
            If isGain Then
                xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/R").Text = CStr(rConfigData.intPRESETGANCool1R)
                xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/G").Text = CStr(rConfigData.intPRESETGANCool1G)
                xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/B").Text = CStr(rConfigData.intPRESETGANCool1B)
            Else
                xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/R").Text = CStr(rConfigData.intPRESETOFFCool1R)
                xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/R").Text = CStr(rConfigData.intPRESETOFFCool1G)
                xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/R").Text = CStr(rConfigData.intPRESETOFFCool1B)
            End If
            
        Case COLORTEMP_STANDARD
            If isGain Then
                xmlDoc.selectSingleNode("/config/PRESETGAN/normal/R").Text = CStr(rConfigData.intPRESETGANNormalR)
                xmlDoc.selectSingleNode("/config/PRESETGAN/normal/G").Text = CStr(rConfigData.intPRESETGANNormalG)
                xmlDoc.selectSingleNode("/config/PRESETGAN/normal/B").Text = CStr(rConfigData.intPRESETGANNormalB)
            Else
                xmlDoc.selectSingleNode("/config/PRESETOFF/normal/R").Text = CStr(rConfigData.intPRESETOFFNormalR)
                xmlDoc.selectSingleNode("/config/PRESETOFF/normal/G").Text = CStr(rConfigData.intPRESETOFFNormalG)
                xmlDoc.selectSingleNode("/config/PRESETOFF/normal/B").Text = CStr(rConfigData.intPRESETOFFNormalB)
            End If
            
        Case COLORTEMP_WARM1
            If isGain Then
                xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/R").Text = CStr(rConfigData.intPRESETGANWarm1R)
                xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/G").Text = CStr(rConfigData.intPRESETGANWarm1G)
                xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/B").Text = CStr(rConfigData.intPRESETGANWarm1B)
            Else
                xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/R").Text = CStr(rConfigData.intPRESETOFFWarm1R)
                xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/G").Text = CStr(rConfigData.intPRESETOFFWarm1G)
                xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/B").Text = CStr(rConfigData.intPRESETOFFWarm1B)
            End If
    End Select
    End If
    
    End
    Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub


Private Sub saveData(strColorTemp As String, HL As Long)

    Select Case strColorTemp
        Case COLORTEMP_COOL1
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

        Case COLORTEMP_STANDARD
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

        Case COLORTEMP_WARM1
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
        Case COLORTEMP_COOL1
            If isGain Then
                rRGB1.cRR = cCOOL1.nColorRR
                rRGB1.cGG = cCOOL1.nColorGG
                rRGB1.cBB = cCOOL1.nColorBB
            Else
                rRGB1.cRR = cFFCOOL1.nColorRR
                rRGB1.cGG = cFFCOOL1.nColorGG
                rRGB1.cBB = cFFCOOL1.nColorBB
            End If
            
        Case COLORTEMP_STANDARD
            If isGain Then
                rRGB1.cRR = cNORMAL.nColorRR
                rRGB1.cGG = cNORMAL.nColorGG
                rRGB1.cBB = cNORMAL.nColorBB
            Else
                rRGB1.cRR = cFFNORMAL.nColorRR
                rRGB1.cGG = cFFNORMAL.nColorGG
                rRGB1.cBB = cFFNORMAL.nColorBB
            End If
            
        Case COLORTEMP_WARM1
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
    Dim sqlstring As String
    Dim cat As New ADOX.Catalog
    Dim tbl As ADOX.Table
    Dim path1 As String
    Dim pstr1 As String
    Dim tabelExist As Boolean

    Set cat = New ADOX.Catalog
    pstr1 = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Data.mdb"
    sqlstring = "select * from [" & gstrCurProjName & "]"
    tabelExist = False
    
    If mBarCode = "" Then
        Exit Sub
    Else
        path1 = Dir(App.Path & "\Data.mdb")
        If path1 = "" Then
            cat.Create pstr1
        End If

        cat.ActiveConnection = pstr1
        For Each tbl In cat.Tables
            If tbl.Name = gstrCurProjName Then
                tabelExist = True
                Exit For
            End If
        Next
        
        If tabelExist = False Then
            Dim tblNew As New Table
            tblNew.Name = gstrCurProjName
            tblNew.Columns.Append "ModelName", adVarWChar, 10
            tblNew.Columns.Append "SerialNO", adVarWChar, 50
            tblNew.Columns.Append "Cool_1x", adInteger
            tblNew.Columns.Append "Cool_1y", adInteger
            tblNew.Columns.Append "Cool_1R", adInteger
            tblNew.Columns.Append "Cool_1G", adInteger
            tblNew.Columns.Append "Cool_1B", adInteger
            tblNew.Columns.Append "Normalx", adInteger
            tblNew.Columns.Append "Normaly", adInteger
            tblNew.Columns.Append "NormalR", adInteger
            tblNew.Columns.Append "NormalG", adInteger
            tblNew.Columns.Append "NormalB", adInteger
            tblNew.Columns.Append "Warm_1x", adInteger
            tblNew.Columns.Append "Warm_1y", adInteger
            tblNew.Columns.Append "Warm_1R", adInteger
            tblNew.Columns.Append "Warm_1G", adInteger
            tblNew.Columns.Append "Warm_1B", adInteger
            tblNew.Columns.Append "OFF_Cool_1R", adInteger
            tblNew.Columns.Append "OFF_Cool_1G", adInteger
            tblNew.Columns.Append "OFF_Cool_1B", adInteger
            tblNew.Columns.Append "OFF_NormalR", adInteger
            tblNew.Columns.Append "OFF_NormalG", adInteger
            tblNew.Columns.Append "OFF_NormalB", adInteger
            tblNew.Columns.Append "OFF_Warm_1R", adInteger
            tblNew.Columns.Append "OFF_Warm_1G", adInteger
            tblNew.Columns.Append "OFF_Warm_1B", adInteger
            tblNew.Columns.Append "Max_Lv", adInteger
            tblNew.Columns.Append "Sepc_Max_Lv", adInteger
            tblNew.Columns.Append "Mark", adVarWChar, 10
            tblNew.Columns.Append "SaveDate", adVarWChar, 10
            tblNew.Columns.Append "SaveTime", adVarWChar, 10
            cat.Tables.Append tblNew
        End If

        FuncOpenSQL (sqlstring)

        rs.AddNew
        
        rs.Fields(0) = gstrCurProjName
        rs.Fields(1) = mBarCode
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
        rs.Fields(27) = glngBlSpecVal
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
    gblnNetConnected = True
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
