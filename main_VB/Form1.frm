VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto Color Temp Adjust System"
   ClientHeight    =   4635
   ClientLeft      =   5865
   ClientTop       =   2625
   ClientWidth     =   10410
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
   ScaleWidth      =   10410
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   495
      Left            =   12480
      TabIndex        =   30
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   375
      Left            =   12600
      TabIndex        =   29
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   375
      Left            =   10560
      TabIndex        =   28
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   255
      Left            =   10560
      TabIndex        =   27
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   10440
      TabIndex        =   26
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   10920
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   10800
      TabIndex        =   24
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   975
      Left            =   11880
      TabIndex        =   21
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
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
      TabIndex        =   8
      Top             =   960
      Width           =   3810
      Begin VB.Label lbColorTempWrong 
         BackStyle       =   0  'Transparent
         Caption         =   "Out Range"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   10
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Left            =   14280
      Top             =   3840
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   12960
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "STOP"
      BeginProperty Font 
         Name            =   "PMingLiU"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13080
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "PMingLiU"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14400
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox CheckStep 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
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
      TabIndex        =   23
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
      TabIndex        =   22
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   10
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
      TabIndex        =   11
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
Dim c14000K As COLORTEMPSPEC
Dim c13000K As COLORTEMPSPEC
Dim c11000K As COLORTEMPSPEC
Dim c9300K As COLORTEMPSPEC
Dim c6500K As COLORTEMPSPEC
Dim cFF13000K As COLORTEMPSPEC
Dim cFF11000K As COLORTEMPSPEC
Dim cFF9300K As COLORTEMPSPEC
Dim cUSERK As COLORTEMPSPEC
Dim rColor As REALCOLOR
Dim Timming, Pattern, Calibrate, MinBrightness As Long
Dim specMaxLV, specMinLV, MaxLV, MinLV As Long
Dim StepTime As Long

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
        txtInput = ""
   Command2.SetFocus
Else
    ShowError_Sys (6)
    GoTo FAIL
End If

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
Call SET_Brightness(50)
DelayMS StepTime

Call SET_Contrast(50)
DelayMS StepTime

Log_Info "###ADJUST COLORTEMP###"
Log_Info "###ADJUST COLORTEMP###"


If IsAdjCool_1 Then
  lbAdjustCOOL_1.BackColor = &H80FFFF
  Result = autoAdjustColorTemperature_Gain(13000, FixG, HighBri)
  If Result = False Then
    ShowError_Sys (1)
    GoTo FAIL
  End If

  lbAdjustCOOL_1.BackColor = &HC0FFC0
End If

If IsAdjNormal Then
    lbAdjustNormal.BackColor = &H80FFFF
  Result = autoAdjustColorTemperature_Gain(11000, FixG, HighBri)
  If Result = False Then
    ShowError_Sys (3)
    GoTo FAIL
  End If

  lbAdjustNormal.BackColor = &HC0FFC0
End If

If IsAdjWarm_1 Then
    lbAdjustWARM_1.BackColor = &H80FFFF
  Result = autoAdjustColorTemperature_Gain(9300, FixG, HighBri)
  If Result = False Then
    ShowError_Sys (4)
    GoTo FAIL
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
         Result = autoAdjustColorTemperature_Offset(13000, FixG, LowBri)
         If Result = False Then
           ShowError_Sys (11)
           GoTo FAIL
         End If
   
         lbAdjustCOOL_1.BackColor = &HC0FFC0
      End If
   End If
   
   If IsAdjNormal Then
     
      If IsAdjsutOffset Then
         lbAdjustNormal.BackColor = &H80FFFF
         Result = autoAdjustColorTemperature_Offset(11000, FixG, LowBri)
         If Result = False Then
           ShowError_Sys (13)
           GoTo FAIL
         End If
    
         lbAdjustNormal.BackColor = &HC0FFC0
      End If
   End If
   
   If IsAdjWarm_1 Then
    
      If IsAdjsutOffset Then
         lbAdjustWARM_1.BackColor = &H80FFFF
         Result = autoAdjustColorTemperature_Offset(9300, FixG, LowBri)
         If Result = False Then
           ShowError_Sys (14)
           GoTo FAIL
         End If
    '
         lbAdjustWARM_1.BackColor = &HC0FFC0
      End If
   End If



End If


  
If IsAdjsutOffset Then Call frmCmbType.ChangePattern(IsWhitePtn)

If IsCheckColorTemp Then
   Label6 = "CHECK"
   If IsAdjCool_1 Then
  
      lbAdjustCOOL_1.BackColor = &H80FFFF
      Result = checkColorAgain(13000, FixG, HighBri)
      If Result = False Then
        ShowError_Sys (1)
        GoTo FAIL
      End If
      lbAdjustCOOL_1.BackColor = &HC0FFC0
   End If
     If IsAdjNormal Then
   
      lbAdjustNormal.BackColor = &H80FFFF
      Result = checkColorAgain(11000, FixG, HighBri)
      If Result = False Then
        ShowError_Sys (3)
        GoTo FAIL
      End If
      lbAdjustNormal.BackColor = &HC0FFC0
   End If
   
     
     If IsAdjWarm_1 Then
    

      lbAdjustWARM_1.BackColor = &H80FFFF
      Result = checkColorAgain(9300, FixG, HighBri)
      If Result = False Then
        ShowError_Sys (4)
        GoTo FAIL
      End If
      lbAdjustWARM_1.BackColor = &HC0FFC0
   End If






End If
  
  
'Save_AllWhiteBlance

DelayMS StepTime

'Save_AllWhiteBlance

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

Label9.Caption = countTime & "S"
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
  RES = setColorTemp(11000, presetData, 0)
End Sub

Private Function autoAdjustColorTemperature_Gain(ColorTemp As Long, FixValue As Long, HighLowMode As Long) As Boolean
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

     SET_RGB_GAN rRGB
     DelayMS StepTime

  showData (1)
  For k = 1 To 50
  If IsStop = True Then GoTo Cancel
  RES = checkColorTemp(rColor, ColorTemp)
  Log_Info "Check colorTemp. RES:" + Str$(RES)
  If RES Then Exit For
  If RES = False Then
     Call adjustColorTemp(FixValue, AdjustSingle, SingleStep, rRGB)

      SET_RGB_GAN rRGB
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
SET_USR_R_GAN rRGB1.cRR
DelayMS StepTime
SET_USR_B_GAN rRGB1.cBB
DelayMS StepTime * 2

     SET_USR_R_OFF rRGB.cRR
     DelayMS StepTime * 2
     SET_USR_G_OFF rRGB.cGG
     DelayMS StepTime * 2
     SET_USR_B_OFF rRGB.cBB
     DelayMS StepTime * 2


  showData (1)

  For k = 1 To 50
  If IsStop = True Then GoTo Cancel
  RES = checkColorTemp(rColor, ColorTemp)
  Log_Info "Check colorTemp. RES:" + Str$(RES)
  If RES Then Exit For
  If RES = False Then
     Call adjustColorTemp(FixValue, AdjustSingle, SingleStep, rRGB)


        SET_USR_R_OFF rRGB.cRR
        DelayMS StepTime * 2
        SET_USR_B_OFF rRGB.cBB
        DelayMS StepTime
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

Private Function checkColorAgain(ColorTemp As Long, FixValue As Long, HighLowMode As Long) As Boolean
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
  RES = checkColorTempTest(rColor, ColorTemp)
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
   xPos = 1515 + (rColor.xx - presetData.xx) * 365 / presetData.xt
   yPos = 1275 - (rColor.yy - presetData.yy) * 385 / presetData.yt
   vPos = 1660 - (rColor.lv - presetData.lv) * 385 / 50
  
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



Private Sub Command1_Click()
IsStop = False
subMainProcesser
                    If IsStop = True Then
                    Exit Sub
                    End If
End Sub

Private Sub Command2_Click()
IsStop = True

txtInput.SetFocus
End Sub





Private Sub Command4_Click()
Dim xx As REALRGB

SET_USR_R_GAN 128 + i
DelayMS StepTime * 3
SET_USR_B_GAN 128 + i
DelayMS StepTime * 2
i = i + 1
End Sub

Private Sub Command5_Click()
SET_COLORTEMP 13000
End Sub

Private Sub Command6_Click()
SET_COLORTEMP 11000
End Sub

Private Sub Command7_Click()
SET_COLORTEMP 9300
End Sub

Private Sub Command8_Click()
'Save_Gain
End Sub

Private Sub Command9_Click()
SET_USR_R_OFF 512 + i
DelayMS StepTime * 3
SET_USR_B_GAN 512 + i
DelayMS StepTime * 2
i = i + 1
End Sub

Private Sub tbAutoADC_Click()
  Form3.Show
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

Private Sub Log_Info(strLog As String)
  CheckStep.Text = CheckStep.Text + strLog + vbCrLf
  CheckStep.SelStart = Len(CheckStep)
  CheckStep.SetFocus
End Sub


Private Sub Form_Load()

i = 0

SetTVCurrentComBaud = 115200

StepTime = IsStepTime
IsStop = False
subInitComPort
subInitInterface
RES = initColorTemp(Timming, Pattern, specMaxLV, specMinLV, Calibrate, MinBrightness, strCurrentModelName)      'InitLPT in dll.

If Timming = 0 Then RES = initColorTemp(Timming, Pattern, specMaxLV, specMinLV, Calibrate, MinBrightness, strCurrentModelName)

DebugFlag = False
Label8 = strCurrentModelName

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
Call Command1_Click
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrExit
  

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
    Case 14000
      c14000K.xx = rColor.xx
      c14000K.yy = rColor.yy
      c14000K.lv = rColor.lv
      c14000K.nColorRR = rRGB.cRR
      c14000K.nColorGG = rRGB.cGG
      c14000K.nColorBB = rRGB.cBB
    Case 13000
    If HL Then
      c13000K.xx = rColor.xx
      c13000K.yy = rColor.yy
      c13000K.lv = rColor.lv
      c13000K.nColorRR = rRGB.cRR
      c13000K.nColorGG = rRGB.cGG
      c13000K.nColorBB = rRGB.cBB
    Else
      cFF13000K.xx = rColor.xx
      cFF13000K.yy = rColor.yy
      cFF13000K.lv = rColor.lv
      cFF13000K.nColorRR = rRGB.cRR
      cFF13000K.nColorGG = rRGB.cGG
      cFF13000K.nColorBB = rRGB.cBB
    End If
    Case 11000
    If HL Then
      c11000K.xx = rColor.xx
      c11000K.yy = rColor.yy
      c11000K.lv = rColor.lv
      c11000K.nColorRR = rRGB.cRR
      c11000K.nColorGG = rRGB.cGG
      c11000K.nColorBB = rRGB.cBB
    Else
      cFF11000K.xx = rColor.xx
      cFF11000K.yy = rColor.yy
      cFF11000K.lv = rColor.lv
      cFF11000K.nColorRR = rRGB.cRR
      cFF11000K.nColorGG = rRGB.cGG
      cFF11000K.nColorBB = rRGB.cBB
    End If
    Case 9300
    If HL Then
      c9300K.xx = rColor.xx
      c9300K.yy = rColor.yy
      c9300K.lv = rColor.lv
      c9300K.nColorRR = rRGB.cRR
      c9300K.nColorGG = rRGB.cGG
      c9300K.nColorBB = rRGB.cBB
    Else
      cFF9300K.xx = rColor.xx
      cFF9300K.yy = rColor.yy
      cFF9300K.lv = rColor.lv
      cFF9300K.nColorRR = rRGB.cRR
      cFF9300K.nColorGG = rRGB.cGG
      cFF9300K.nColorBB = rRGB.cBB
    End If
    Case 6500
      c6500K.xx = rColor.xx
      c6500K.yy = rColor.yy
      c6500K.lv = rColor.lv
      c6500K.nColorRR = rRGB.cRR
      c6500K.nColorGG = rRGB.cGG
      c6500K.nColorBB = rRGB.cBB
    Case 1000
      cUSERK.xx = rColor.xx
      cUSERK.yy = rColor.yy
      cUSERK.lv = rColor.lv
      cUSERK.nColorRR = rRGB.cRR
      cUSERK.nColorGG = rRGB.cGG
      cUSERK.nColorBB = rRGB.cBB

  End Select
  
  
End Sub

Private Sub LoadData(ColorTemp As Long)
  Select Case ColorTemp
    Case 14000
      rRGB1.cRR = c14000K.nColorRR

      rRGB1.cBB = c14000K.nColorBB
    Case 13000

      rRGB1.cRR = c13000K.nColorRR

      rRGB1.cBB = c13000K.nColorBB

    Case 11000

      rRGB1.cRR = c11000K.nColorRR

      rRGB1.cBB = c11000K.nColorBB

    Case 9300
      rRGB1.cRR = c9300K.nColorRR

      rRGB1.cBB = c9300K.nColorBB

    Case 6500
      rRGB1.cRR = c6500K.nColorRR

      rRGB1.cBB = c6500K.nColorBB
    Case 1000

      cUSERK.nColorRR = rRGB.cRR

      cUSERK.nColorBB = rRGB.cBB

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

  rs.Fields(2) = c13000K.xx
  rs.Fields(3) = c13000K.yy
  rs.Fields(4) = c13000K.lv
  rs.Fields(5) = c13000K.nColorRR
  rs.Fields(6) = c13000K.nColorGG
  rs.Fields(7) = c13000K.nColorBB
  rs.Fields(8) = c11000K.xx
  rs.Fields(9) = c11000K.yy
  rs.Fields(10) = c11000K.lv
  rs.Fields(11) = c11000K.nColorRR
  rs.Fields(12) = c11000K.nColorGG
  rs.Fields(13) = c11000K.nColorBB
  rs.Fields(14) = c9300K.xx
  rs.Fields(15) = c9300K.yy
  rs.Fields(16) = c9300K.lv
  rs.Fields(17) = c9300K.nColorRR
  rs.Fields(18) = c9300K.nColorGG
  rs.Fields(19) = c9300K.nColorBB
  
  rs.Fields(20) = cFF13000K.xx
  rs.Fields(21) = cFF13000K.yy
  rs.Fields(22) = cFF13000K.lv
  rs.Fields(23) = cFF13000K.nColorRR
  rs.Fields(24) = cFF13000K.nColorGG
  rs.Fields(25) = cFF13000K.nColorBB
  rs.Fields(26) = cFF11000K.xx
  rs.Fields(27) = cFF11000K.yy
  rs.Fields(28) = cFF11000K.lv
  rs.Fields(29) = cFF11000K.nColorRR
  rs.Fields(30) = cFF11000K.nColorGG
  rs.Fields(31) = cFF11000K.nColorBB
  rs.Fields(32) = cFF9300K.xx
  rs.Fields(33) = cFF9300K.yy
  rs.Fields(34) = cFF9300K.lv
  rs.Fields(35) = cFF9300K.nColorRR
  rs.Fields(36) = cFF9300K.nColorGG
  rs.Fields(37) = cFF9300K.nColorBB

  rs.Fields(38) = MinLV
  rs.Fields(39) = MaxLV

  rs.Fields(40) = cmdMark
  rs.Fields(41) = Date
  rs.Fields(42) = Time
  
  rs.Fields(43) = c14000K.xx
  rs.Fields(44) = c14000K.yy
  rs.Fields(45) = c14000K.lv
  rs.Fields(46) = c14000K.nColorRR
  rs.Fields(47) = c14000K.nColorGG
  rs.Fields(48) = c14000K.nColorBB
  rs.Fields(49) = c6500K.xx
  rs.Fields(50) = c6500K.yy
  rs.Fields(51) = c6500K.lv
  rs.Fields(52) = c6500K.nColorRR
  rs.Fields(53) = c6500K.nColorGG
  rs.Fields(54) = c6500K.nColorBB
  rs.Fields(55) = cUSERK.xx
  rs.Fields(56) = cUSERK.yy
  rs.Fields(57) = cUSERK.lv
  rs.Fields(58) = cUSERK.nColorRR
  rs.Fields(59) = cUSERK.nColorGG
  rs.Fields(60) = cUSERK.nColorBB
  
  rs.Update

 Set cn = Nothing
 Set rs = Nothing
 sqlstring = ""
End If
End Sub

