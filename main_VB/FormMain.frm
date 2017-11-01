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
   Begin MSWinsockLib.Winsock tcpServer 
      Left            =   10560
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
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
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label lbSpec 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SPEC"
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
      Top             =   4020
      Width           =   2535
   End
   Begin VB.Label lbMeasure 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Measure"
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
      Top             =   3510
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
   Begin VB.Label lbAdjustWarm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Warm"
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
      Top             =   2490
      Width           =   2535
   End
   Begin VB.Label lbAdjustStandard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Standard"
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
      Top             =   1980
      Width           =   2535
   End
   Begin VB.Label lbAdjustCool 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cool"
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
         Caption         =   "Common Settings"
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
Dim cCool As COLORTEMPSPEC
Dim cStandard As COLORTEMPSPEC
Dim cWarm As COLORTEMPSPEC
Dim cFFCool As COLORTEMPSPEC
Dim cFFStandard As COLORTEMPSPEC
Dim cFFWarm As COLORTEMPSPEC
Dim rColor As REALCOLOR
Dim lvLastChk As Long
Dim resCodeForAdjustColorTemp As Long
Dim clsProtocal As Protocal
Dim clsCANTVProtocal As CANTVProtocal
Dim clsLetvProtocal As LetvProtocal
Dim clsLetvCurvedProtocal As LetvCurvedProtocal
Dim clsLetvMST6M60 As LetvMST6M60
Dim clsHaierProtocal As HaierProtocal
Dim clsKONKAProtocal As KONKAProtocal

Dim ivpg As IVPGCtrl

Private rRGB As REALRGB
Private rRGB1 As REALRGB

Private mudtPreColorData As COLORTEMPSPEC
Private mAdjGainAgainCool As Integer
Private mAdjGainAgainStandard As Integer
Private mAdjGainAgainWarm As Integer
Private mCntTime As Long
Private mTitle As String
Private mBrand As String
Private mBarCode As String

Private WithEvents Obj As VPGCtrl.VPGCtrl
Attribute Obj.VB_VarHelpID = -1


Private Sub Form_Load()
    vbFunc.Caption = TXTFun
    vbConCA310.Caption = TXTConnectCA
    tbDisConnectastro.Caption = TXTDisConnectCA
    vbSet.Caption = TXTSet
    vbSetSPEC.Caption = TXTSetSpec
    vbDescription.Caption = TXTDiscription
    vbAbout.Caption = TXTAbout
    Label6.Caption = TXTINITIAL
    checkResult.Caption = TXTChkResult
    gblnStop = False
    txtInput.Enabled = True
    
    Me.Caption = TXTTitle & " V" & App.Major & "." & App.Minor & "." & App.Revision
    mTitle = Me.Caption
    SubInit

    mBrand = Split(gstrCurProjName, DELIMITER)(0)
    
    If UCase(mBrand) = "CAN" Then    'CANTV
        Set clsCANTVProtocal = New CANTVProtocal
        Set clsProtocal = clsCANTVProtocal
        PictureBrand.Picture = LoadPicture(App.Path & "\Resources\CANTV.bmp")
    ElseIf UCase(mBrand) = "HAIER" Then    'Haier
        Set clsHaierProtocal = New HaierProtocal
        Set clsProtocal = clsHaierProtocal
        PictureBrand.Picture = LoadPicture(App.Path & "\Resources\Haier.bmp")
    ElseIf UCase(mBrand) = "KONKA" Then    'KONKA
        Set clsKONKAProtocal = New KONKAProtocal
        Set clsProtocal = clsKONKAProtocal
        PictureBrand.Picture = LoadPicture(App.Path & "\Resources\KONKA.bmp")
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
    
    RES = ColorTInit(gstrCurProjName, App.Path)
End Sub

Private Sub Form_Unload(Cancel As Integer)
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
    
    If gEnumCommMode = modeNetServer Then
        If tcpServer.State <> sckClosed Then
            tcpServer.Close
        End If
    End If
  
    Call ColorTDeInit

    Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
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

Private Sub tbDisConnectastro_Click()
    If gblnCaConnected Then
        ObjCa.RemoteMode = 0
    End If
End Sub

Private Sub Timer1_Timer()
    mCntTime = mCntTime + 1
    lbTimer.Caption = CStr(mCntTime) & "s"
End Sub

Private Sub SubInitNetClient()
    gblnNetConnected = False
    With tcpClient
        .Protocol = sckTCPProtocol
        ' IMPORTANT: be sure to change the RemoteHost
        ' value to the name of your computer.
        .RemoteHost = REMOTE_HOST
        .RemotePort = REMOTE_PORT
    End With
End Sub

Private Sub tcpClient_Connect()
    'Success to connect the TV.
    gblnNetConnected = True
End Sub

Private Sub SubInitNetServer()
    If tcpServer.State <> sckListening Then
        With tcpServer
            .Protocol = sckTCPProtocol
            .LocalPort = PORT_FOR_KONKA
            .Listen
        End With
        txtInput.Enabled = False
        CheckStep.Text = "Please enter factory menu and select [Auto White Balance]"
    End If
End Sub

Private Sub tcpServer_ConnectionRequest(ByVal requestID As Long)
    ' Check if the control's State is closed. If not,
    ' close the connection before accepting the new
    ' connection.
    If tcpServer.State <> sckClosed Then
        tcpServer.Close
        ' Accept the request with the requestID
        ' parameter.
        tcpServer.Accept requestID
        txtInput.Enabled = True
        txtInput.Text = ""
        txtInput.SetFocus
        CheckStep.Text = "Connect TV successfully."
    End If
End Sub

Public Sub SubInit()
    LoadConfigData

    gstrCurProjName = gudtConfigData.strModel
    gEnumCommMode = gudtConfigData.CommMode
    gintCurComBaud = gudtConfigData.strComBaud
    gintCurComId = gudtConfigData.intComID
    glngI2cClockRate = gudtConfigData.lngI2cClockRate
    gstrTvInputSrc = gudtConfigData.strInputSource
    gintTvInputSrcPort = CInt(Right(gstrTvInputSrc, 1))
    gstrTvInputSrc = Left(gstrTvInputSrc, Len(gstrTvInputSrc) - 1)
    glngDelayTime = gudtConfigData.lngDelayMs
    glngCaChannel = gudtConfigData.intChannelNum
    gintBarCodeLen = gudtConfigData.intBarCodeLen
    glngBlSpecVal = gudtConfigData.intLvSpec
    gstrVPGModel = gudtConfigData.strVPGModel
    gstrVPGTiming = gudtConfigData.strVPGTiming
    gstrVPG100IRE = gudtConfigData.strVPG100IRE
    gstrVPG80IRE = gudtConfigData.strVPG80IRE
    gstrVPG20IRE = gudtConfigData.strVPG20IRE
    gblnEnableCool = gudtConfigData.bolEnableCool
    gblnEnableStandard = gudtConfigData.bolEnableStandard
    gblnEnableWarm = gudtConfigData.bolEnableWarm
    gblnChkColorTemp = gudtConfigData.bolEnableChkColor
    gblnAdjOffset = gudtConfigData.bolEnableAdjOffset
    gstrChipSet = gudtConfigData.strChipSet
    
    If gEnumCommMode = modeUART Then
        SubInitComPort
    ElseIf gEnumCommMode = modeNetClient Then
        SubInitNetClient
    ElseIf gEnumCommMode = modeNetServer Then
        SubInitNetServer
    End If

    txtInput.Text = ""
    lbModelName.Caption = Split(gstrCurProjName, DELIMITER)(1)
    
    If gblnEnableCool = True Then lbAdjustCool.ForeColor = &H80000008
    If gblnEnableStandard = True Then lbAdjustStandard.ForeColor = &H80000008
    If gblnEnableWarm = True Then lbAdjustWarm.ForeColor = &H80000008

    If gblnEnableCool = False Then lbAdjustCool.ForeColor = &HC0C0C0
    If gblnEnableStandard = False Then lbAdjustStandard.ForeColor = &HC0C0C0
    If gblnEnableWarm = False Then lbAdjustWarm.ForeColor = &HC0C0C0
    
    SubInitVPG
    SubDelayMs 200
    
    Call SubVPGTiming(gstrVPGTiming)
End Sub

Private Sub SubInitComPort()
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

Private Sub SubInitVPG()
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

Private Sub SubVPGTiming(Tim As String)
    Dim bNo(1) As Byte
    
    bNo(0) = (CInt(Tim) And &HFF00) \ 256
    bNo(1) = CInt(Tim) And &HFF

    ivpg.ExecuteCmd VPG_CMD_CM_DOWNLOAD, VPG_SCMD_SCM_CTL_RUNTIM, bNo, False
End Sub

Private Sub SubVPGPattern(Ptn As String)
    Dim bNo(1) As Byte
    
    bNo(0) = (CInt(Ptn) And &HFF00) \ 256
    bNo(1) = CInt(Ptn) And &HFF

    ivpg.RunKey (VPG_KEY_CKEY_OUT)
    ivpg.ExecuteCmd VPG_CMD_CM_DOWNLOAD, VPG_SCMD_SCM_CTL_RUNPTN, bNo, False
End Sub

Private Sub Obj_OnChangedConnectState(ByVal bIsConnected As Boolean)
    If bIsConnected = False Then
        Me.Caption = mTitle & " [Chroma " & gstrVPGModel & " Disconnected]"
    Else
        Me.Caption = mTitle
    End If
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

            If gEnumCommMode = modeUART Then
                If MSComm1.PortOpen = False Then
                    MSComm1.PortOpen = True
                End If
                SubRun
            ElseIf gEnumCommMode = modeNetClient Then
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
            ElseIf gEnumCommMode = modeI2c Then
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
            ElseIf gEnumCommMode = modeNetServer Then
                If tcpServer.State = sckConnected Then
                    SubRun
                Else
                    SubLogInfo "TCP Server state is: " + CStr(tcpServer.State)
                    If tcpServer.State <> sckClosed Then
                        tcpServer.Close
                    End If
                    SubInitNetServer
                    SubLogInfo "Please enter factory menu and select [Auto White Balance]"
                End If
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

Private Sub SubRun()
    On Error GoTo ErrExit
    Dim Result As Boolean

    SubInitBeforeRun

    If gblnStop = True Then
        Exit Sub
    End If

    If gblnCaConnected = False Then
        MsgBox TXTCaDisconnectHint, vbOKOnly + vbInformation, "warning"
        SubConfigAfterRun

        Exit Sub
    End If

    gblnStop = False
    checkResult.BackColor = &H80FFFF
    checkResult.Caption = TXTRun
    checkResult.ForeColor = &HC0&
    CheckStep.Text = ""

    lbAdjustCool.BackColor = &H8000000F
    lbAdjustStandard.BackColor = &H8000000F
    lbAdjustWarm.BackColor = &H8000000F

    Picture1.Cls
    lbColorTempWrong.Visible = False

    Set ObjMemory = ObjCa.Memory
    ObjMemory.ChannelNO = glngCaChannel

    SubLogInfo "Start..."
    Call SubVPGPattern(gstrVPG80IRE)

    clsProtocal.EnterFacMode
    Call clsProtocal.SwitchInputSource(gstrTvInputSrc, gintTvInputSrcPort)
    'Call clsProtocal.ResetPicMode
    Call clsProtocal.SetBrightness(50)
    Call clsProtocal.SetContrast(50)
    Call clsProtocal.SetBacklight(100)
    SubLogInfo "Set Brightness: 50, Set Contrast: 50, Set Backlight: 100"

    Label6.Caption = "WHITE"

ADJUST_GAIN_AGAIN_COOL:
    If gblnEnableCool Then
        lbAdjustCool.BackColor = &H80FFFF
        Result = FuncAdjRGBGain(COLORTEMP_COOL, ADJMODE_3)
  
        If Result = False Then
            ShowError_Sys (1)
            GoTo FAIL
        Else
            Call clsProtocal.SaveWBDataToAllSrc(gstrTvInputSrc, gintTvInputSrcPort)
        End If

        SubSaveLogInFile "[Time]White Cool: " & lbTimer.Caption
        lbAdjustCool.BackColor = &HC0FFC0
        
        If mAdjGainAgainCool > 0 Then
            GoTo CHECK_COOL
        End If
    End If

ADJUST_GAIN_AGAIN_STANDARD:
    If gblnEnableStandard Then
        lbAdjustStandard.BackColor = &H80FFFF
        Result = FuncAdjRGBGain(COLORTEMP_STANDARD, ADJMODE_3)

        If Result = False Then
            ShowError_Sys (2)
            GoTo FAIL
        Else
            Call clsProtocal.SaveWBDataToAllSrc(gstrTvInputSrc, gintTvInputSrcPort)
        End If

        SubSaveLogInFile "[Time]White Standard: " & lbTimer.Caption
        lbAdjustStandard.BackColor = &HC0FFC0
        
        If mAdjGainAgainStandard > 0 Then
            GoTo CHECK_STANDARD
        End If
    End If

ADJUST_GAIN_AGAIN_WARM:
    If gblnEnableWarm Then
        lbAdjustWarm.BackColor = &H80FFFF
        Result = FuncAdjRGBGain(COLORTEMP_WARM, ADJMODE_3)

        If Result = False Then
            ShowError_Sys (3)
            GoTo FAIL
        Else
            Call clsProtocal.SaveWBDataToAllSrc(gstrTvInputSrc, gintTvInputSrcPort)
        End If

        SubSaveLogInFile "[Time]White Warm: " & lbTimer.Caption
        lbAdjustWarm.BackColor = &HC0FFC0
        
        If mAdjGainAgainWarm > 0 Then
            GoTo CHECK_WARM
        End If
    End If

    If gblnAdjOffset Then
        Label6.Caption = TXTGrey

        Call SubVPGPattern(gstrVPG20IRE)

        If gblnEnableCool Then
            lbAdjustCool.BackColor = &H80FFFF
            Result = FuncAdjRGBOffset(COLORTEMP_COOL)
                
            If Result = False Then
                ShowError_Sys (5)
                GoTo FAIL
            Else
                Call clsProtocal.SaveWBDataToAllSrc(gstrTvInputSrc, gintTvInputSrcPort)
            End If
            
            SubSaveLogInFile "[Time]Grey Cool: " & lbTimer.Caption
            lbAdjustCool.BackColor = &HC0FFC0
        End If
   
        If gblnEnableStandard Then
            lbAdjustStandard.BackColor = &H80FFFF
            Result = FuncAdjRGBOffset(COLORTEMP_STANDARD)

            If Result = False Then
                ShowError_Sys (6)
                GoTo FAIL
            Else
                Call clsProtocal.SaveWBDataToAllSrc(gstrTvInputSrc, gintTvInputSrcPort)
            End If

            SubSaveLogInFile "[Time]Grey Standard: " & lbTimer.Caption
            lbAdjustStandard.BackColor = &HC0FFC0
        End If
   
        If gblnEnableWarm Then
            lbAdjustWarm.BackColor = &H80FFFF
            Result = FuncAdjRGBOffset(COLORTEMP_WARM)
                
            If Result = False Then
                ShowError_Sys (7)
                GoTo FAIL
            Else
                Call clsProtocal.SaveWBDataToAllSrc(gstrTvInputSrc, gintTvInputSrcPort)
            End If

            SubSaveLogInFile "[Time]Grey Warm: " & lbTimer.Caption
            lbAdjustWarm.BackColor = &HC0FFC0
        End If
    End If

    If gblnChkColorTemp Then
        If gblnAdjOffset Then
            Call SubVPGPattern(gstrVPG80IRE)
        End If

CHECK_COOL:
        If gblnEnableCool Then
            Label6.Caption = TXTChk
            lbAdjustCool.BackColor = &H80FFFF
            Result = FuncChkColorAgain(COLORTEMP_COOL)

            If Result = False Then
                ShowError_Sys (1)

                If mAdjGainAgainCool > 0 Then
                    GoTo FAIL
                End If
                
                mAdjGainAgainCool = mAdjGainAgainCool + 1
                
                GoTo ADJUST_GAIN_AGAIN_COOL
            Else
                mAdjGainAgainCool = 0
            End If
      
            lbAdjustCool.BackColor = &HC0FFC0
        End If

CHECK_STANDARD:
        If gblnEnableStandard Then
            Label6.Caption = TXTChk
            lbAdjustStandard.BackColor = &H80FFFF
            Result = FuncChkColorAgain(COLORTEMP_STANDARD)

            If Result = False Then
                ShowError_Sys (2)

                If mAdjGainAgainStandard > 0 Then
                    GoTo FAIL
                End If
    
                mAdjGainAgainStandard = mAdjGainAgainStandard + 1

                GoTo ADJUST_GAIN_AGAIN_STANDARD
            Else
                mAdjGainAgainStandard = 0
            End If
    
            lbAdjustStandard.BackColor = &HC0FFC0
        End If

CHECK_WARM:
        If gblnEnableWarm Then
            Label6.Caption = TXTChk
            lbAdjustWarm.BackColor = &H80FFFF
            Result = FuncChkColorAgain(COLORTEMP_WARM)

            If Result = False Then
                ShowError_Sys (3)
                
                If mAdjGainAgainWarm > 0 Then
                    GoTo FAIL
                End If
    
                mAdjGainAgainWarm = mAdjGainAgainWarm + 1
                
                GoTo ADJUST_GAIN_AGAIN_WARM
            Else
                mAdjGainAgainWarm = 0
            End If

            lbAdjustWarm.BackColor = &HC0FFC0
        End If
    End If
    
    If (gstrChipSet = "T111" Or gstrChipSet = "Hi3751") Then
        Call clsProtocal.SelColorTemp(COLORTEMP_STANDARD, gstrTvInputSrc, gintTvInputSrcPort)
        SubLogInfo "Set color temp to Standard."
        
        ObjCa.Measure
        lvLastChk = CLng(ObjProbe.lv)
        SubLogInfo "lv = " + CStr(lvLastChk)
        SubShowData (LASTSTEP)
    Else
        'Last check:
        'Cool, 100% white pattern, brightness = 100, contrast = 100
        'Check Lv and save x, y, lv
        Call SubVPGPattern(gstrVPG100IRE)

        Call clsProtocal.SetBrightness(100)
        SubLogInfo "Set brightness to 100"

        Call clsProtocal.SetContrast(100)
        SubLogInfo "Set contrast to 100"

        If UCase(mBrand) = "LETV" Then
            Call clsProtocal.SelColorTemp(COLORTEMP_COOL, gstrTvInputSrc, gintTvInputSrcPort)
            SubLogInfo "Set color temp to Cool."
        Else
            Call clsProtocal.SelColorTemp(COLORTEMP_STANDARD, gstrTvInputSrc, gintTvInputSrcPort)
            SubLogInfo "Set color temp to Standard."
        End If

        ObjCa.Measure
        lvLastChk = CLng(ObjProbe.lv)
        SubLogInfo "lv = " + CStr(lvLastChk)
        SubShowData (LASTSTEP)

        Call clsProtocal.SetBrightness(50)
        Call clsProtocal.SetContrast(50)
        SubLogInfo "Set both brightness and contrast to 50."
    
        'clsProtocal.ResetPicMode
        clsProtocal.ChannelPreset

        If lvLastChk <= glngBlSpecVal Then
            ShowError_Sys (8)
            GoTo FAIL
        End If
    End If

PASS:
    clsProtocal.ExitFacMode

    Call SubSaveDataToDB(TXTPass)

    CheckStep = CheckStep + "TEST ALL PASS"
    CheckStep.SelStart = Len(CheckStep)
    checkResult.ForeColor = &HC000&
    checkResult.Caption = TXTPass
    checkResult.BackColor = &HFF00&
    checkResult.ForeColor = &HC00000
    
    Label6.Caption = TXTPass
    
    Call SubConfigAfterRun

    Exit Sub

FAIL:
    clsProtocal.ExitFacMode

    Call SubSaveDataToDB(TXTFail)

    CheckStep.SelStart = Len(CheckStep)
    checkResult.BackColor = &HFF&
    checkResult.ForeColor = &H808080
    checkResult.Caption = TXTFail
    checkResult.ForeColor = &H0&
    checkResult.ForeColor = &HFFFF&
    
    Label6.Caption = TXTFail

    Call SubConfigAfterRun

    Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub SubInitBeforeRun()
    mCntTime = 0
    lbTimer.Caption = "0s"
    Timer1.Enabled = True

    txtInput.Enabled = False
    'mBarCode = ""
    mAdjGainAgainCool = 0
    mAdjGainAgainStandard = 0
    mAdjGainAgainWarm = 0
End Sub

Private Sub SubConfigAfterRun()
    Timer1.Enabled = False
    
    SubSaveLogInFile "[Time]Total: " & lbTimer.Caption & vbCrLf
    
    mAdjGainAgainCool = 0
    mAdjGainAgainStandard = 0
    mAdjGainAgainWarm = 0

    If gEnumCommMode = modeNetServer Then
        If tcpServer.State <> sckClosed Then
            tcpServer.Close
        End If
        SubInitNetServer
    Else
        txtInput.Enabled = True
        txtInput.Text = ""
        txtInput.SetFocus
        
        If gEnumCommMode = modeNetClient Then
            If tcpClient.State <> sckClosed Then
                gblnNetConnected = False
                tcpClient.Close
            End If
        End If
    End If
End Sub

Private Sub ShowError_Sys(t As Integer)
    Dim s As String
    
    s = "Unknown"

    Select Case t
        Case 1
            s = TXTGainCoolWrong
        Case 2
            s = TXTGainStandardWrong
        Case 3
            s = TXTGainWarmWrong
        Case 4
            s = "LAB_SN:" + mBarCode + "(End)  Len:" + str$(gintBarCodeLen) + vbCrLf + TXTSNLenWrong
        Case 5
            s = TXTOffsetCoolWrong
        Case 6
            s = TXTOffsetStandardWrong
        Case 7
            s = TXTOffsetWarmWrong
        Case 8
            s = TXTLvTooLow
    End Select

    CheckStep.Text = CheckStep.Text + TXTErrCode + str$(t) + vbCrLf + s + vbCrLf
    CheckStep.SelStart = Len(CheckStep)
End Sub

Private Function FuncAdjRGBGain(strColorTemp As String, adjustVal As Long) As Boolean
    Dim i, j As Integer
    
    SubLogInfo "====SelColorTemp " + strColorTemp + "===="
    Call clsProtocal.SelColorTemp(strColorTemp, gstrTvInputSrc, gintTvInputSrcPort)

    SubLogInfo "========Adjust " & strColorTemp & "========"

    ' Set RGB Offset
    Call ColorTSetSpec(strColorTemp, mudtPreColorData, ADJMODE_OFFSET)
        
    rRGB.cRR = mudtPreColorData.nColorRR
    rRGB.cGG = mudtPreColorData.nColorGG
    rRGB.cBB = mudtPreColorData.nColorBB

    Call SubUpdateRGB(strColorTemp, 0)

    Call LoadData(strColorTemp, 0)
    If UCase(gstrChipSet) = "MST6M60" Then
        Call clsProtocal.SetRGBOffset(rRGB1.cRR * 8, rRGB1.cGG * 8, rRGB1.cBB * 8)
    Else
        Call clsProtocal.SetRGBOffset(rRGB1.cRR, rRGB1.cGG, rRGB1.cBB)
    End If

    For i = 1 To 2
        Call ColorTSetSpec(strColorTemp, mudtPreColorData, ADJMODE_GAIN)
        'SubDelayMs 200
        
        SubLogInfo "Init current colorTemp. RES:" + str$(RES)
        rRGB.cRR = mudtPreColorData.nColorRR
        rRGB.cGG = mudtPreColorData.nColorGG
        rRGB.cBB = mudtPreColorData.nColorBB
        
        Label1 = CStr(mudtPreColorData.xx)
        Label3 = CStr(mudtPreColorData.yy)

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
            SubLogInfo "SPEC: x = " & CStr(mudtPreColorData.xx) & " y = " & CStr(mudtPreColorData.yy)
            SubLogInfo "Tol: x = " & CStr(mudtPreColorData.xt) & " y =  " & CStr(mudtPreColorData.yt)

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
        Call SubUpdateRGB(strColorTemp, ADJMODE_GAIN)
        SubLogInfo "Save current data of " & strColorTemp & "."
        FuncAdjRGBGain = True
    Else
        FuncAdjRGBGain = False
    End If

End Function

Private Function FuncAdjRGBOffset(strColorTemp As String) As Boolean
    Dim i, j As Integer

    SubLogInfo "====SelColorTemp " + strColorTemp + "===="
    Call clsProtocal.SelColorTemp(strColorTemp, gstrTvInputSrc, gintTvInputSrcPort)

    SubLogInfo "========Adjust " & strColorTemp & "========"
  
    For i = 1 To 2
        Call ColorTSetSpec(strColorTemp, mudtPreColorData, ADJMODE_OFFSET)
        'SubDelayMs 200
        SubLogInfo "Init current colorTemp. RES:" + str$(RES)
        rRGB.cRR = mudtPreColorData.nColorRR
        rRGB.cGG = mudtPreColorData.nColorGG
        rRGB.cBB = mudtPreColorData.nColorBB
  
        Label1 = CStr(mudtPreColorData.xx)
        Label3 = CStr(mudtPreColorData.yy)

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
        Call SubUpdateRGB(strColorTemp, ADJMODE_OFFSET)
        SubLogInfo "Save current data of " & strColorTemp & "."
        FuncAdjRGBOffset = True
    Else
        FuncAdjRGBOffset = False
    End If

End Function

Private Function FuncChkColorAgain(strColorTemp As String) As Boolean
    Dim i As Integer

    SubLogInfo "====SelColorTemp " + strColorTemp + "===="
    Call clsProtocal.SelColorTemp(strColorTemp, gstrTvInputSrc, gintTvInputSrcPort)

    SubLogInfo "========Check " & strColorTemp & "========"
  
    For i = 1 To 2
        Call ColorTSetSpec(strColorTemp, mudtPreColorData, ADJMODE_GAIN)
        'SubDelayMs 200
        SubLogInfo "Init current colorTemp. RES:" + str$(RES)

        Label1 = CStr(mudtPreColorData.xx)
        Label3 = CStr(mudtPreColorData.yy)

        SubShowData (5)

        If gblnStop = True Then GoTo Cancel

        RES = ColorTChk(rColor, strColorTemp)
        SubLogInfo "Check colorTemp. RES:" + str$(RES)

        If RES = 3 Then Exit For
    Next i
  
Cancel:
    If RES = 3 Then
        FuncChkColorAgain = True
    Else
        FuncChkColorAgain = False
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
    If mudtPreColorData.xt = 0 Then
        mudtPreColorData.xt = 30
    End If
    If mudtPreColorData.yt = 0 Then
        mudtPreColorData.yt = 30
    End If
    xPos = 1515 + (rColor.xx - mudtPreColorData.xx) * 365 / mudtPreColorData.xt
    yPos = 1275 - (rColor.yy - mudtPreColorData.yy) * 385 / mudtPreColorData.yt

    If step = LASTSTEP Then
        vPos = 1660 - (rColor.lv - glngBlSpecVal) * 385 / 50
    Else
        vPos = 1660 - (rColor.lv - mudtPreColorData.lv) * 385 / 50
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
        If Abs(rColor.xx - mudtPreColorData.xx) <= mudtPreColorData.xt And Abs(rColor.yy - mudtPreColorData.yy) <= mudtPreColorData.yt Then
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
        If rColor.lv > mudtPreColorData.lv Then
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

Private Sub SubUpdateRGB(strColorTemp As String, HL As Long)

    Select Case strColorTemp
        Case COLORTEMP_COOL
            If HL Then
                cCool.xx = rColor.xx
                cCool.yy = rColor.yy
                cCool.lv = rColor.lv
                cCool.nColorRR = rRGB.cRR
                cCool.nColorGG = rRGB.cGG
                cCool.nColorBB = rRGB.cBB
            Else
                cFFCool.xx = rColor.xx
                cFFCool.yy = rColor.yy
                cFFCool.lv = rColor.lv
                cFFCool.nColorRR = rRGB.cRR
                cFFCool.nColorGG = rRGB.cGG
                cFFCool.nColorBB = rRGB.cBB
            End If

        Case COLORTEMP_STANDARD
            If HL Then
                cStandard.xx = rColor.xx
                cStandard.yy = rColor.yy
                cStandard.lv = rColor.lv
                cStandard.nColorRR = rRGB.cRR
                cStandard.nColorGG = rRGB.cGG
                cStandard.nColorBB = rRGB.cBB
            Else
                cFFStandard.xx = rColor.xx
                cFFStandard.yy = rColor.yy
                cFFStandard.lv = rColor.lv
                cFFStandard.nColorRR = rRGB.cRR
                cFFStandard.nColorGG = rRGB.cGG
                cFFStandard.nColorBB = rRGB.cBB
            End If

        Case COLORTEMP_WARM
            If HL Then
                cWarm.xx = rColor.xx
                cWarm.yy = rColor.yy
                cWarm.lv = rColor.lv
                cWarm.nColorRR = rRGB.cRR
                cWarm.nColorGG = rRGB.cGG
                cWarm.nColorBB = rRGB.cBB
            Else
                cFFWarm.xx = rColor.xx
                cFFWarm.yy = rColor.yy
                cFFWarm.lv = rColor.lv
                cFFWarm.nColorRR = rRGB.cRR
                cFFWarm.nColorGG = rRGB.cGG
                cFFWarm.nColorBB = rRGB.cBB
            End If
    End Select
  
End Sub

Private Sub LoadData(strColorTemp As String, isGain As Boolean)
    Select Case strColorTemp
        Case COLORTEMP_COOL
            If isGain Then
                rRGB1.cRR = cCool.nColorRR
                rRGB1.cGG = cCool.nColorGG
                rRGB1.cBB = cCool.nColorBB
            Else
                rRGB1.cRR = cFFCool.nColorRR
                rRGB1.cGG = cFFCool.nColorGG
                rRGB1.cBB = cFFCool.nColorBB
            End If
            
        Case COLORTEMP_STANDARD
            If isGain Then
                rRGB1.cRR = cStandard.nColorRR
                rRGB1.cGG = cStandard.nColorGG
                rRGB1.cBB = cStandard.nColorBB
            Else
                rRGB1.cRR = cFFStandard.nColorRR
                rRGB1.cGG = cFFStandard.nColorGG
                rRGB1.cBB = cFFStandard.nColorBB
            End If
            
        Case COLORTEMP_WARM
            If isGain Then
                rRGB1.cRR = cWarm.nColorRR
                rRGB1.cGG = cWarm.nColorGG
                rRGB1.cBB = cWarm.nColorBB
            Else
                rRGB1.cRR = cFFWarm.nColorRR
                rRGB1.cGG = cFFWarm.nColorGG
                rRGB1.cBB = cFFWarm.nColorBB
            End If
    End Select
End Sub

Private Sub SubSaveDataToDB(strMark As String)
    Dim sqlstring As String
    Dim cat As New ADOX.Catalog
    Dim tbl As ADOX.Table
    Dim path1 As String
    Dim pstr1 As String
    Dim tabelExist As Boolean

    Set cat = New ADOX.Catalog
    pstr1 = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\" & MDB_FILE_NAME
    sqlstring = "select * from [" & gstrCurProjName & "]"
    tabelExist = False
    
    If mBarCode = "" Then
        Exit Sub
    Else
        path1 = Dir(App.Path & "\" & MDB_FILE_NAME)
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
            tblNew.Columns.Append "ModelName", adVarWChar, 30
            tblNew.Columns.Append "SerialNO", adVarWChar, 50
            tblNew.Columns.Append "Coolx", adInteger
            tblNew.Columns.Append "Cooly", adInteger
            tblNew.Columns.Append "CoolR", adInteger
            tblNew.Columns.Append "CoolG", adInteger
            tblNew.Columns.Append "CoolB", adInteger
            tblNew.Columns.Append "Standardx", adInteger
            tblNew.Columns.Append "Standardy", adInteger
            tblNew.Columns.Append "StandardR", adInteger
            tblNew.Columns.Append "StandardG", adInteger
            tblNew.Columns.Append "StandardB", adInteger
            tblNew.Columns.Append "Warmx", adInteger
            tblNew.Columns.Append "Warmy", adInteger
            tblNew.Columns.Append "WarmR", adInteger
            tblNew.Columns.Append "WarmG", adInteger
            tblNew.Columns.Append "WarmB", adInteger
            tblNew.Columns.Append "OFF_CoolR", adInteger
            tblNew.Columns.Append "OFF_CoolG", adInteger
            tblNew.Columns.Append "OFF_CoolB", adInteger
            tblNew.Columns.Append "OFF_StandardR", adInteger
            tblNew.Columns.Append "OFF_StandardG", adInteger
            tblNew.Columns.Append "OFF_StandardB", adInteger
            tblNew.Columns.Append "OFF_WarmR", adInteger
            tblNew.Columns.Append "OFF_WarmG", adInteger
            tblNew.Columns.Append "OFF_WarmB", adInteger
            tblNew.Columns.Append "Max_Lv", adInteger
            tblNew.Columns.Append "Sepc_Max_Lv", adInteger
            tblNew.Columns.Append "Mark", adVarWChar, 10
            tblNew.Columns.Append "SaveDate", adVarWChar, 20
            tblNew.Columns.Append "SaveTime", adVarWChar, 20
            cat.Tables.Append tblNew
        End If
        
        FuncOpenSQL (sqlstring)
        
        rs.AddNew
        
        rs.Fields(0) = gstrCurProjName
        rs.Fields(1) = mBarCode
        rs.Fields(2) = cCool.xx
        rs.Fields(3) = cCool.yy
        rs.Fields(4) = cCool.nColorRR
        rs.Fields(5) = cCool.nColorGG
        rs.Fields(6) = cCool.nColorBB
        rs.Fields(7) = cStandard.xx
        rs.Fields(8) = cStandard.yy
        rs.Fields(9) = cStandard.nColorRR
        rs.Fields(10) = cStandard.nColorGG
        rs.Fields(11) = cStandard.nColorBB
        rs.Fields(12) = cWarm.xx
        rs.Fields(13) = cWarm.yy
        rs.Fields(14) = cWarm.nColorRR
        rs.Fields(15) = cWarm.nColorGG
        rs.Fields(16) = cWarm.nColorBB
        rs.Fields(17) = cFFCool.nColorRR
        rs.Fields(18) = cFFCool.nColorGG
        rs.Fields(19) = cFFCool.nColorBB
        rs.Fields(20) = cFFStandard.nColorRR
        rs.Fields(21) = cFFStandard.nColorGG
        rs.Fields(22) = cFFStandard.nColorBB
        rs.Fields(23) = cFFWarm.nColorRR
        rs.Fields(24) = cFFWarm.nColorGG
        rs.Fields(25) = cFFWarm.nColorBB
        rs.Fields(26) = lvLastChk
        rs.Fields(27) = glngBlSpecVal
        rs.Fields(28) = strMark
        rs.Fields(29) = Date
        rs.Fields(30) = Time
        
        rs.Update

        Set cn = Nothing
        Set rs = Nothing
        sqlstring = ""
    End If
End Sub
