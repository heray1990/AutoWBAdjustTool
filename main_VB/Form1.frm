VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto Color Temp Adjust System"
   ClientHeight    =   4635
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
   ScaleHeight     =   4635
   ScaleWidth      =   10335
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      DrawWidth       =   3
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   2640
      Picture         =   "Form1.frx":57E2
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
      Text            =   "Form1.frx":248DC
      Top             =   960
      Width           =   3805
   End
   Begin VB.TextBox txtInput 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      TabIndex        =   1
      Text            =   "123456789"
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
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sampl1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "0S"
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
      Caption         =   "WHITE"
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
      Begin VB.Menu tbDebugMode 
         Caption         =   "DebugMode(&M)"
      End
      Begin VB.Menu tbAutoADC 
         Caption         =   "ReadMAC(&A)"
         Shortcut        =   {F8}
      End
      Begin VB.Menu vbConChroma 
         Caption         =   "ConnectChroma"
      End
      Begin VB.Menu vbEXIT 
         Caption         =   "EXIT"
      End
   End
   Begin VB.Menu vbSET 
      Caption         =   "Setting"
      Begin VB.Menu tbSetComPort 
         Caption         =   "Set ComPort(&P)"
      End
      Begin VB.Menu vbSetSPEC 
         Caption         =   "Set SPEC"
      End
   End
   Begin VB.Menu vbDescription 
      Caption         =   "Description"
      Begin VB.Menu vbAbout 
         Caption         =   "About"
         Shortcut        =   {F2}
      End
      Begin VB.Menu vbHELP 
         Caption         =   "Help"
         Shortcut        =   {F1}
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
Dim Timming, Pattern, Calibrate, MinBrightness As Long
Dim specMaxLV, specMinLV, MaxLV, MinLV As Long
Dim StepTime As Long
Dim resCodeForAdjustColorTemp As Long

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
    
    ENTER_FAC_MODE
    DelayMS StepTime
    
    SEL_INPUT_HDMI1_FOR_WB
    DelayMS StepTime
    
On Error GoTo ErrExit

    If IsCa210ok = False Then
        MsgBox "CA210 disconnected,Please click'Connect'->'Connect CA210'to do operation!", vbOKOnly + vbInformation, "warning"
        txtInput.Text = ""
        txtInput.SetFocus
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
    Label6 = "WHITE"

    Picture1.Cls
    lbColorTempWrong.Visible = False
    
    If IsAdjCool_1 = False Then lbAdjustCOOL_1.ForeColor = &HC0C0C0
    If IsAdjCool_2 = False Then lbAdjustCOOL_2.ForeColor = &HC0C0C0
    If IsAdjNormal = False Then lbAdjustNormal.ForeColor = &HC0C0C0
    If IsAdjWarm_1 = False Then lbAdjustWARM_1.ForeColor = &HC0C0C0
    If IsAdjWarm_2 = False Then lbAdjustWARM_2.ForeColor = &HC0C0C0

    Set ObjMemory = ObjCa.Memory
    ObjMemory.ChannelNO = IsCa210Channel
    
    If IsAdjsutOffset Then
        Call frmCmbType.ChangePattern(IsWhitePtn)
        DelayMS 200
    End If
    strBuff = ""

    Log_Info "###INITIAL USER###"
    Log_Info "###INITIAL USER###"

    Log_Info "###ADJUST COLORTEMP###"
    Log_Info "###ADJUST COLORTEMP###"

    If IsAdjCool_1 Then
        lbAdjustCOOL_1.BackColor = &H80FFFF
        Result = autoAdjustColorTemperature_Gain(valColorTempCool1, adjustMode3, HighBri)
  
        If Result = False Then
            ShowError_Sys (1)
            GoTo FAIL
        Else
            SAVE_WB_DATA_TO_ALL_SRC
            DelayMS StepTime * 2
        End If

        lbAdjustCOOL_1.BackColor = &HC0FFC0
    End If

    If IsAdjNormal Then
        lbAdjustNormal.BackColor = &H80FFFF
        Result = autoAdjustColorTemperature_Gain(valColorTempNormal, adjustMode3, HighBri)

        If Result = False Then
            ShowError_Sys (3)
            GoTo FAIL
        Else
            SAVE_WB_DATA_TO_ALL_SRC
            DelayMS StepTime * 2
        End If

        lbAdjustNormal.BackColor = &HC0FFC0
    End If

    If IsAdjWarm_1 Then
        lbAdjustWARM_1.BackColor = &H80FFFF
        Result = autoAdjustColorTemperature_Gain(valColorTempWarm1, adjustMode3, HighBri)

        If Result = False Then
            ShowError_Sys (4)
            GoTo FAIL
        Else
            SAVE_WB_DATA_TO_ALL_SRC
            DelayMS StepTime * 2
        End If

        lbAdjustWARM_1.BackColor = &HC0FFC0
    End If
  
    If IsSendOffset Then
        Label6 = "GREY"

        If IsAdjsutOffset Then
            Call frmCmbType.ChangePattern("109")
            DelayMS 800
        End If

        DelayMS StepTime

        If IsAdjCool_1 Then
            If IsAdjsutOffset Then
                lbAdjustCOOL_1.BackColor = &H80FFFF
                Result = autoAdjustColorTemperature_Offset(valColorTempCool1, FixG, LowBri)
                
                If Result = False Then
                    ShowError_Sys (11)
                    GoTo FAIL
                Else
                    SAVE_WB_DATA_TO_ALL_SRC
                    DelayMS StepTime * 2
                End If
   
                lbAdjustCOOL_1.BackColor = &HC0FFC0
            End If
        End If
   
        If IsAdjNormal Then
            If IsAdjsutOffset Then
                lbAdjustNormal.BackColor = &H80FFFF
                Result = autoAdjustColorTemperature_Offset(valColorTempNormal, FixG, LowBri)

                If Result = False Then
                    ShowError_Sys (13)
                    GoTo FAIL
                Else
                    SAVE_WB_DATA_TO_ALL_SRC
                    DelayMS StepTime * 2
                End If
    
                lbAdjustNormal.BackColor = &HC0FFC0
            End If
        End If
   
        If IsAdjWarm_1 Then
            If IsAdjsutOffset Then
                lbAdjustWARM_1.BackColor = &H80FFFF
                Result = autoAdjustColorTemperature_Offset(valColorTempWarm1, FixG, LowBri)
                
                If Result = False Then
                    ShowError_Sys (14)
                    GoTo FAIL
                Else
                    SAVE_WB_DATA_TO_ALL_SRC
                    DelayMS StepTime * 2
                End If

                lbAdjustWARM_1.BackColor = &HC0FFC0
            End If
        End If

    End If
  
    If IsAdjsutOffset Then Call frmCmbType.ChangePattern(IsWhitePtn)

    If IsCheckColorTemp Then
        Label6 = "CHECK"

        If IsAdjCool_1 Then
            lbAdjustCOOL_1.BackColor = &H80FFFF
            Result = checkColorAgain(valColorTempCool1, adjustMode3, HighBri)

            If Result = False Then
                ShowError_Sys (1)
                GoTo FAIL
            End If
      
            lbAdjustCOOL_1.BackColor = &HC0FFC0
        End If
     
        If IsAdjNormal Then
            lbAdjustNormal.BackColor = &H80FFFF
            Result = checkColorAgain(valColorTempNormal, adjustMode3, HighBri)
      
            If Result = False Then
                ShowError_Sys (3)
                GoTo FAIL
            End If
    
            lbAdjustNormal.BackColor = &HC0FFC0
        End If
     
        If IsAdjWarm_1 Then
            lbAdjustWARM_1.BackColor = &H80FFFF
            Result = checkColorAgain(valColorTempWarm1, adjustMode3, HighBri)

            If Result = False Then
                ShowError_Sys (4)
                GoTo FAIL
            End If

            lbAdjustWARM_1.BackColor = &HC0FFC0
        End If

    End If
    
    'Last check:
    'Cool, 100% white pattern, brightness = 100, contrast = 100
    'Check Lv and save x, y, lv
    Call frmCmbType.ChangePattern("101")
    DelayMS StepTime
    
    SET_BRIGHTNESS 100
    DelayMS StepTime
    Log_Info "Set brightness to 100"
    
    SET_CONTRAST 100
    DelayMS StepTime
    Log_Info "Set contrast to 100"
    
    SET_COLORTEMP valColorTempCool1
    DelayMS StepTime
    Log_Info "Set color temp to cool1"

    DelayMS StepTime
    ObjCa.Measure
    rColorLastChk.xx = CLng(ObjProbe.sx * 10000)
    rColorLastChk.yy = CLng(ObjProbe.sy * 10000)
    rColorLastChk.lv = CLng(ObjProbe.lv)
    
    Log_Info "x = " + Str$(rColorLastChk.xx) + ", y = " + Str$(rColorLastChk.yy) + ", lv = " + Str$(rColorLastChk.lv)

    If rColorLastChk.lv < specMinLV Then
        Log_Info "亮度不在规格！"
        GoTo FAIL
    End If

    EXIT_FAC_MODE
    DelayMS StepTime

    Call saveALLcData

PASS:
    CheckStep.ForeColor = &H80000008
    CheckStep.BackColor = &HC0FFC0
    checkResult.ForeColor = &HC000&
    checkResult.Caption = "PASS"
    DelayMS StepTime
    CheckStep = CheckStep + strSerialNo + vbCrLf
    CheckStep = CheckStep + strSerialNo + vbCrLf
    CheckStep = CheckStep + "TEST ALL PASS"
    CheckStep.SelStart = Len(CheckStep)
    CheckStep.SetFocus
    Call subInitAfterRunning
    checkResult.BackColor = &HFF00&
    checkResult.ForeColor = &HC00000

    Exit Sub

FAIL:
    If IsAdjsutOffset Then Call frmCmbType.ChangePattern(IsWhitePtn)
    
    CheckStep.SelStart = Len(CheckStep)
    CheckStep.SetFocus
    Call subInitAfterRunning
    checkResult.BackColor = &HFF&
    CheckStep.BackColor = &HFFFF&
    checkResult.ForeColor = &H808080
    checkResult.Caption = "FAIL"
    DelayMS 1900
    checkResult.ForeColor = &H0&
    DelayMS 300
    checkResult.ForeColor = &HFFFF&

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
    countTime = Timer
    IsSNWriteSuccess = True
    strSerialNo = ""
End Sub

Private Sub subInitAfterRunning()
    countTime = CLng(Timer - countTime)

    Label9.Caption = countTime & "s"
    IsSNWriteSuccess = False

    txtInput.Text = ""
    txtInput.SetFocus
End Sub

Private Function subJudgeTheSNIsAvailable() As Boolean
    If strSerialNo = "" Or Len(strSerialNo) <> IsBarcodeLen Then
        CheckStep.Text = ""
        CheckStep.Text = CheckStep.Text + "Please confirm the SN again?" + vbCrLf
        txtInput.Text = ""
        txtInput.SetFocus
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
            s = "LAB_SN:" + strSerialNo + "(End)  Len:" + Str$(IsBarcodeLen) + vbCrLf + "Barcode SerialNumber is Wrong"
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
            s = ""
    End Select

    CheckStep.ForeColor = &HFF&
    CheckStep.Text = CheckStep.Text + "Error Code:" + Str$(t) + vbCrLf + s + vbCrLf
    CheckStep.SelStart = Len(CheckStep)
    CheckStep.SetFocus
End Sub

Private Sub Command3_Click()
  RES = setColorTemp(valColorTempNormal, presetData, 0)
End Sub

Private Function autoAdjustColorTemperature_Gain(ColorTemp As Long, adjustVal As Long, HighLowMode As Long) As Boolean
    Dim i, j, k As Integer
  
    Log_Info "========Adjust " + Str$(ColorTemp) + "K========"
  
    For j = 1 To 2
        SET_COLORTEMP ColorTemp
        DelayMS StepTime
  
        Call setColorTemp(ColorTemp, presetData, HighLowMode)
        DelayMS StepTime
        
        Log_Info "Init current colorTemp. RES:" + Str$(RES)
        rRGB.cRR = presetData.nColorRR
        rRGB.cGG = presetData.nColorGG
        rRGB.cBB = presetData.nColorBB
        
        Label1 = Str$(presetData.xx)
        Label3 = Str$(presetData.yy)

        'SET_RGB_GAN rRGB
        SET_R_GAN rRGB.cRR
        DelayMS StepTime
        
        SET_G_GAN rRGB.cGG
        DelayMS StepTime
        
        SET_B_GAN rRGB.cBB
        DelayMS StepTime

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
 
                'SET_RGB_GAN rRGB
                SET_R_GAN rRGB.cRR
                DelayMS StepTime
                
                SET_G_GAN rRGB.cGG
                DelayMS StepTime
                
                SET_B_GAN rRGB.cBB
                DelayMS StepTime

                showData (2)
            End If
  
            'DelayMS StepTime
        Next k
  
        If RES Then Exit For
        
        'DelayMS StepTime
    Next j

Cancel:
    If RES Then
        'Save_Gain
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
        SET_COLORTEMP ColorTemp
        DelayMS StepTime

        Call setColorTemp(ColorTemp, presetData, HighLowMode)
        DelayMS StepTime
        Log_Info "Init current colorTemp. RES:" + Str$(RES)
        rRGB.cRR = presetData.nColorRR
        rRGB.cGG = presetData.nColorGG
        rRGB.cBB = presetData.nColorBB
  
        Label1 = Str$(presetData.xx)
        Label3 = Str$(presetData.yy)

        Call LoadData(ColorTemp)

        SET_R_GAN rRGB1.cRR
        DelayMS StepTime

        SET_B_GAN rRGB1.cBB
        DelayMS StepTime

        SET_R_OFF rRGB.cRR
        DelayMS StepTime
     
        SET_G_OFF rRGB.cGG
        DelayMS StepTime
     
        SET_B_OFF rRGB.cBB
        DelayMS StepTime

        showData (1)

        For k = 1 To 50
            If IsStop = True Then GoTo Cancel
            
            RES = checkColorTemp(rColor, ColorTemp)
            Log_Info "Check colorTemp. RES:" + Str$(RES)

            If RES Then Exit For
            If RES = False Then
                Call adjustColorTempOffset(FixValue, AdjustSingle, SingleStep, rRGB)

                SET_R_OFF rRGB.cRR
                DelayMS StepTime

                SET_B_OFF rRGB.cBB
                DelayMS StepTime

                showData (2)
            End If

            DelayMS 200
        Next k

        If RES Then Exit For

        DelayMS StepTime
    Next j
 
Cancel:
    If RES Then
        'Save_Gain
   
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
        SET_COLORTEMP ColorTemp
  
        Call setColorTemp(ColorTemp, presetData, HighLowMode)
        DelayMS StepTime
        Log_Info "Init current colorTemp. RES:" + Str$(RES)

        Label1 = Str$(presetData.xx)
        Label3 = Str$(presetData.yy)

        showData (1)

        If IsStop = True Then GoTo Cancel
        'RES = checkColorTempTest(rColor, ColorTemp)
        RES = checkColorTemp(rColor, ColorTemp)
        Log_Info "Check colorTemp. RES:" + Str$(RES)

        If RES Then Exit For

        DelayMS StepTime
    Next j
  
Cancel:
    If RES Then
        checkColorAgain = True
    Else
        checkColorAgain = False
    End If

End Function


Private Sub showData(step As Integer)
On Error Resume Next
    Dim xPos, yPos, vPos As Long
    
    DelayMS StepTime
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
    vPos = 1660 - (rColor.lv - presetData.lv) * 385 / 50

    'In dx-dy axis, 360 is the distance from left edge of white rectangle to the left of Picture1.
    'In dx-dy axis, 2660 is the distance from right edge of white rectangle to the left of Picture1.
    'In dx-dy axis, 80 is the distance from top edge of white rectangle to the top of Picture1.
    'In dx-dy axis, 2660 is the distance from bottom edge of white rectangle to the top of Picture1.
    If xPos < 360 Then xPos = 360
    If xPos > 2660 Then xPos = 2660
    If yPos < 80 Then yPos = 80
    If yPos > 2480 Then yPos = 2480
   
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

    'In lv axis, 3060 is the distance from left edge of white rectangle to the left of Picture1.
    'In lv axis, 3390 is the distance from right edge of white rectangle to the left of Picture1.
    If rColor.lv > presetData.lv Then
        Picture1.Line (3060, vPos)-(3390, vPos), &H30FF30
    Else
        Picture1.Line (3060, vPos)-(3390, vPos), &HFF&
    End If
 
    Log_Info "_x/y/Lv: " + Str$(rColor.xx) + " / " + Str$(rColor.yy) + " / " + Str$(rColor.lv)

    If Label6 <> "CHECK" Then Log_Info "_R/G/B:  " + Str$(rRGB.cRR) + " / " + Str$(rRGB.cGG) + " / " + Str$(rRGB.cBB)

    Label_x = Str$(rColor.xx)
    Label_y = Str$(rColor.yy)
    Label_Lv = Str$(rColor.lv)
    DelayMS 30

    If DebugFlag Then
        DelayMS 2000
    End If
End Sub

Private Sub tbDebugMode_Click()
    DebugFlag = True
End Sub

Private Sub tbDisConnectastro_Click()
    ObjCa.RemoteMode = 0
End Sub

Private Sub tbSetComPort_Click()
    Form2.Show
End Sub

Private Sub vbConChroma_Click()
    frmCmbType.Show
End Sub

Private Sub vbEXIT_Click()
    Unload Me
    End
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
    SetTVCurrentComBaud = 115200
    StepTime = IsStepTime
    IsStop = False
    
    subInitComPort
    subInitInterface
    
    Label8 = strCurrentModelName
    
    RES = initColorTemp(Timming, Pattern, specMaxLV, specMinLV, Calibrate, MinBrightness, strCurrentModelName, App.path)      'InitLPT in dll.

    If Timming = 0 Then
        RES = initColorTemp(Timming, Pattern, specMaxLV, specMinLV, Calibrate, MinBrightness, strCurrentModelName, App.path)
    End If

    DebugFlag = False

    If IsAdjCool_1 = False Then lbAdjustCOOL_1.ForeColor = &HC0C0C0
    If IsAdjCool_2 = False Then lbAdjustCOOL_2.ForeColor = &HC0C0C0
    If IsAdjNormal = False Then lbAdjustNormal.ForeColor = &HC0C0C0
    If IsAdjWarm_1 = False Then lbAdjustWARM_1.ForeColor = &HC0C0C0
    If IsAdjWarm_2 = False Then lbAdjustWARM_2.ForeColor = &HC0C0C0
End Sub

Private Sub subInitInterface()
    txtInput.Text = ""
End Sub

Private Sub subInitComPort()
    sqlstring = "select * from CommonTable where Mark='ATS'"
    Executesql (sqlstring)
    
    If rs.EOF = False Then
        SetTVCurrentComID = rs("ComID")
    Else
        MsgBox "Read Data Error,Please Check Your Database!", vbOKOnly + vbInformation, "Warning!"
        End
    End If

    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""
    
    ComInit
End Sub


Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        IsStop = False
        Call subMainProcesser
        
        If IsStop = True Then
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrExit

    IsStop = True
    If (IsCa210ok = True) Then
        ObjCa.RemoteMode = 0
    End If
  
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
  
    Call DeinitColorTemp(strCurrentModelName)
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

Private Sub LoadData(ColorTemp As Long)
    Select Case ColorTemp
        Case valColorTempCool1
            rRGB1.cRR = c12000K.nColorRR
            rRGB1.cBB = c12000K.nColorBB
            
        Case valColorTempNormal
            rRGB1.cRR = c10000K.nColorRR
            rRGB1.cBB = c10000K.nColorBB
            
        Case valColorTempWarm1
            rRGB1.cRR = c6500K.nColorRR
            rRGB1.cBB = c6500K.nColorBB
    End Select
End Sub

Private Sub saveALLcData()

    Dim cmdFTP As String
    Dim cmdMark As String
    Dim OffsetRGB As String
    
    If strSerialNo = "" Then
        Exit Sub
    Else
        cmdMark = "Y"
        sqlstring = "select * from DataRecord"
        Executesql (sqlstring)
        rs.AddNew

        rs.Fields(0) = strCurrentModelName
        rs.Fields(1) = strSerialNo

        rs.Fields(2) = c12000K.xx
        rs.Fields(3) = c12000K.yy
        rs.Fields(4) = c12000K.lv
        rs.Fields(5) = c12000K.nColorRR
        rs.Fields(6) = c12000K.nColorGG
        rs.Fields(7) = c12000K.nColorBB
        rs.Fields(8) = c10000K.xx
        rs.Fields(9) = c10000K.yy
        rs.Fields(10) = c10000K.lv
        rs.Fields(11) = c10000K.nColorRR
        rs.Fields(12) = c10000K.nColorGG
        rs.Fields(13) = c10000K.nColorBB
        rs.Fields(14) = c6500K.xx
        rs.Fields(15) = c6500K.yy
        rs.Fields(16) = c6500K.lv
        rs.Fields(17) = c6500K.nColorRR
        rs.Fields(18) = c6500K.nColorGG
        rs.Fields(19) = c6500K.nColorBB
  
        rs.Fields(20) = cFF12000K.xx
        rs.Fields(21) = cFF12000K.yy
        rs.Fields(22) = cFF12000K.lv
        rs.Fields(23) = cFF12000K.nColorRR
        rs.Fields(24) = cFF12000K.nColorGG
        rs.Fields(25) = cFF12000K.nColorBB
        rs.Fields(26) = cFF10000K.xx
        rs.Fields(27) = cFF10000K.yy
        rs.Fields(28) = cFF10000K.lv
        rs.Fields(29) = cFF10000K.nColorRR
        rs.Fields(30) = cFF10000K.nColorGG
        rs.Fields(31) = cFF10000K.nColorBB
        rs.Fields(32) = cFF6500K.xx
        rs.Fields(33) = cFF6500K.yy
        rs.Fields(34) = cFF6500K.lv
        rs.Fields(35) = cFF6500K.nColorRR
        rs.Fields(36) = cFF6500K.nColorGG
        rs.Fields(37) = cFF6500K.nColorBB

        rs.Fields(38) = MinLV
        rs.Fields(39) = MaxLV

        rs.Fields(40) = rColorLastChk.xx
        rs.Fields(41) = rColorLastChk.yy
        rs.Fields(42) = rColorLastChk.lv
        rs.Fields(43) = specMinLV
        rs.Fields(44) = cmdMark
        rs.Fields(45) = Date
        rs.Fields(46) = Time
  
        rs.Update

        Set cn = Nothing
        Set rs = Nothing
        sqlstring = ""
    End If

End Sub

