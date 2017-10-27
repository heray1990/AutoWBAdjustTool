VERSION 5.00
Begin VB.Form FormSettings 
   Caption         =   "Common Settings"
   ClientHeight    =   6630
   ClientLeft      =   6435
   ClientTop       =   3210
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormSettings.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   6630
   ScaleWidth      =   5055
   Begin VB.Frame Frame7 
      Caption         =   "Chroma"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   32
      Top             =   3000
      Width           =   2400
      Begin VB.ComboBox cmbChromaModel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FormSettings.frx":1DF72
         Left            =   1200
         List            =   "FormSettings.frx":1DFA0
         TabIndex        =   41
         Text            =   "22294"
         Top             =   280
         Width           =   1000
      End
      Begin VB.TextBox txtChromaTiming 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   40
         Text            =   "1"
         Top             =   720
         Width           =   1000
      End
      Begin VB.Frame Frame8 
         Caption         =   "Pattern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   2175
         Begin VB.TextBox txt100IRE 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1150
            TabIndex        =   36
            Text            =   "1"
            Top             =   240
            Width           =   800
         End
         Begin VB.TextBox txt80IRE 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1150
            TabIndex        =   35
            Text            =   "1"
            Top             =   600
            Width           =   800
         End
         Begin VB.TextBox txt20IRE 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1150
            TabIndex        =   34
            Text            =   "1"
            Top             =   960
            Width           =   800
         End
         Begin VB.Label lb100IRE 
            Caption         =   "100IRE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   270
            Width           =   795
         End
         Begin VB.Label lb80IRE 
            Caption         =   "80IRE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   630
            Width           =   795
         End
         Begin VB.Label lb20IRE 
            Caption         =   "20IRE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   990
            Width           =   795
         End
      End
      Begin VB.Label lbChromaModel 
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   43
         Top             =   330
         Width           =   900
      End
      Begin VB.Label lbChromaTiming 
         Caption         =   "Timing:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   42
         Top             =   750
         Width           =   900
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "I2C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   2550
      TabIndex        =   29
      Top             =   3240
      Width           =   2400
      Begin VB.ComboBox cmbI2cClockRate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FormSettings.frx":1E00D
         Left            =   1200
         List            =   "FormSettings.frx":1E00F
         TabIndex        =   30
         Text            =   "50KHz"
         Top             =   300
         Width           =   1000
      End
      Begin VB.Label lbClockRate 
         Caption         =   "Clock Rate:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   31
         Top             =   330
         Width           =   1140
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Serial Port"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2550
      TabIndex        =   23
      Top             =   2040
      Width           =   2400
      Begin VB.ComboBox cmbComID 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FormSettings.frx":1E011
         Left            =   1200
         List            =   "FormSettings.frx":1E013
         TabIndex        =   25
         Text            =   "COM1"
         Top             =   300
         Width           =   1000
      End
      Begin VB.ComboBox cmbComBaud 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FormSettings.frx":1E015
         Left            =   1200
         List            =   "FormSettings.frx":1E017
         TabIndex        =   24
         Text            =   "9600"
         Top             =   660
         Width           =   1000
      End
      Begin VB.Label lbComId 
         Caption         =   "ComID:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   27
         Top             =   330
         Width           =   900
      End
      Begin VB.Label lbComBaud 
         Caption         =   "ComBaud:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   26
         Top             =   700
         Width           =   900
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Communication Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2550
      TabIndex        =   20
      Top             =   840
      Width           =   2400
      Begin VB.OptionButton optNetServer 
         Caption         =   "NetServer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   1100
         TabIndex        =   44
         Top             =   720
         Width           =   1125
      End
      Begin VB.OptionButton optI2c 
         Caption         =   "I2C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   800
      End
      Begin VB.OptionButton optUart 
         Caption         =   "UART"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   800
      End
      Begin VB.OptionButton optNetClient 
         Caption         =   "NetClient"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   1100
         TabIndex        =   21
         Top             =   360
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Common Setting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2550
      TabIndex        =   8
      Top             =   4080
      Width           =   2400
      Begin VB.TextBox txtLvSpec 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   19
         Text            =   "280"
         Top             =   1350
         Width           =   1000
      End
      Begin VB.ComboBox cmbInputSource 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FormSettings.frx":1E019
         Left            =   1200
         List            =   "FormSettings.frx":1E029
         TabIndex        =   15
         Text            =   "HDMI1"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox txtSNLen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   11
         Text            =   "1"
         Top             =   650
         Width           =   1000
      End
      Begin VB.TextBox txtDelay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   9
         Text            =   "500"
         Top             =   300
         Width           =   1000
      End
      Begin VB.Label lbLvSpec 
         Caption         =   "Lv Spec:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   18
         Top             =   1380
         Width           =   900
      End
      Begin VB.Label lbInputSrc 
         Caption         =   "TV Source:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   16
         Top             =   1030
         Width           =   900
      End
      Begin VB.Label lbSnLen 
         Caption         =   "SN_Len:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   12
         Top             =   680
         Width           =   900
      End
      Begin VB.Label lbDelayMs 
         Caption         =   "Delay(ms):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   10
         Top             =   330
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "CA310/CA210"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   120
      TabIndex        =   5
      Top             =   5760
      Width           =   2400
      Begin VB.TextBox txtChannel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   7
         Text            =   "1"
         Top             =   300
         Width           =   1000
      End
      Begin VB.Label lbChannelNum 
         Caption         =   "Channel:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   6
         Top             =   330
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2400
      Begin VB.CheckBox Check7 
         Alignment       =   1  'Right Justify
         Caption         =   "Adjust Offset"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   14
         Top             =   1750
         Value           =   1  'Checked
         Width           =   1900
      End
      Begin VB.CheckBox Check6 
         Alignment       =   1  'Right Justify
         Caption         =   "Check Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   13
         Top             =   1400
         Value           =   1  'Checked
         Width           =   1900
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "Cool"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   0
         Top             =   350
         Value           =   1  'Checked
         Width           =   1900
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   "Standard"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   1
         Top             =   700
         Value           =   1  'Checked
         Width           =   1900
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "Warm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   2
         Top             =   1050
         Value           =   1  'Checked
         Width           =   1900
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      TabIndex        =   3
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "FormSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_Load()
    Dim i As Integer

    Label1.Caption = gstrCurProjName
    
    txtChannel.Text = CStr(glngCaChannel)
    txtSNLen.Text = CStr(gintBarCodeLen)
    txtLvSpec.Text = CStr(glngBlSpecVal)
    cmbInputSource.Text = gstrTvInputSrc & CStr(gintTvInputSrcPort)
    txtDelay.Text = glngDelayTime

    cmbComBaud.Text = CStr(gintCurComBaud)
    cmbComID.Text = "COM" & CStr(gintCurComId)
    For i = 1 To 20
        cmbComID.AddItem "COM" & i
    Next i

    cmbComBaud.AddItem "9600"
    cmbComBaud.AddItem "19200"
    cmbComBaud.AddItem "38400"
    cmbComBaud.AddItem "57600"
    cmbComBaud.AddItem "115200"
    
    cmbChromaModel.Text = gstrVPGModel
    txtChromaTiming.Text = gstrVPGTiming
    txt100IRE.Text = gstrVPG100IRE
    txt80IRE.Text = gstrVPG80IRE
    txt20IRE.Text = gstrVPG20IRE
    cmbI2cClockRate.Text = glngI2cClockRate & "KHz"

    If gEnumCommMode = modeUART Then
        optUart.Value = True
        optNetClient.Value = False
        optI2c.Value = False
        optNetServer.Value = False
        cmbComBaud.Enabled = True
        cmbComID.Enabled = True
        cmbI2cClockRate.Enabled = False
    ElseIf gEnumCommMode = modeNetClient Then
        optUart.Value = False
        optNetClient.Value = True
        optI2c.Value = False
        optNetServer.Value = False
        cmbComBaud.Enabled = False
        cmbComID.Enabled = False
        cmbI2cClockRate.Enabled = False
    ElseIf gEnumCommMode = modeI2c Then
        optUart.Value = False
        optNetClient.Value = False
        optI2c.Value = True
        optNetServer.Value = False
        cmbComBaud.Enabled = False
        cmbComID.Enabled = False
        cmbI2cClockRate.Enabled = True
    ElseIf gEnumCommMode = modeNetServer Then
        optUart.Value = False
        optNetClient.Value = False
        optI2c.Value = False
        optNetServer.Value = True
        cmbComBaud.Enabled = False
        cmbComID.Enabled = False
        cmbI2cClockRate.Enabled = False
    End If

    If gblnEnableCool Then
        Check2.Value = 1
    Else
        Check2.Value = 0
    End If

    If gblnEnableStandard Then
        Check3.Value = 1
    Else
        Check3.Value = 0
    End If

    If gblnEnableWarm Then
        Check4.Value = 1
    Else
        Check4.Value = 0
    End If

    If gblnChkColorTemp Then
        Check6.Value = 1
    Else
        Check6.Value = 0
    End If

    If gblnAdjOffset Then
        Check7.Value = 1
    Else
        Check7.Value = 0
    End If
    
End Sub

Private Sub Command1_Click()
    If Check2.Value = 1 Then gudtConfigData.bolEnableCool = True
    If Check2.Value = 0 Then gudtConfigData.bolEnableCool = False
    If Check3.Value = 1 Then gudtConfigData.bolEnableStandard = True
    If Check3.Value = 0 Then gudtConfigData.bolEnableStandard = False
    If Check4.Value = 1 Then gudtConfigData.bolEnableWarm = True
    If Check4.Value = 0 Then gudtConfigData.bolEnableWarm = False
    If Check6.Value = 1 Then gudtConfigData.bolEnableChkColor = True
    If Check6.Value = 0 Then gudtConfigData.bolEnableChkColor = False
    If Check7.Value = 1 Then gudtConfigData.bolEnableAdjOffset = True
    If Check7.Value = 0 Then gudtConfigData.bolEnableAdjOffset = False
    
    If optUart.Value = True Then
        gudtConfigData.CommMode = modeUART
    ElseIf optNetClient.Value = True Then
        gudtConfigData.CommMode = modeNetClient
    ElseIf optI2c.Value = True Then
        gudtConfigData.CommMode = modeI2c
    ElseIf optNetServer.Value = True Then
        gudtConfigData.CommMode = modeNetServer
    Else
        gudtConfigData.CommMode = modeUART
    End If

    gudtConfigData.strComBaud = cmbComBaud.Text
    gudtConfigData.intComID = val(Replace(cmbComID.Text, "COM", ""))
    gudtConfigData.lngI2cClockRate = val(Replace(cmbI2cClockRate.Text, "KHz", ""))
    gudtConfigData.strInputSource = cmbInputSource.Text
    gudtConfigData.lngDelayMs = val(txtDelay.Text)
    gudtConfigData.intChannelNum = val(txtChannel.Text)
    gudtConfigData.intBarCodeLen = val(txtSNLen.Text)
    gudtConfigData.intLvSpec = val(txtLvSpec.Text)
    gudtConfigData.strVPGModel = cmbChromaModel.Text
    gudtConfigData.strVPGTiming = txtChromaTiming.Text
    gudtConfigData.strVPG100IRE = txt100IRE.Text
    gudtConfigData.strVPG80IRE = txt80IRE.Text
    gudtConfigData.strVPG20IRE = txt20IRE.Text
    
    SaveConfigData
    

    Unload Me
    
    FormMain.SubInit
    FormMain.Show
End Sub

Private Sub optI2c_Click()
    cmbComBaud.Enabled = False
    cmbComID.Enabled = False
    cmbI2cClockRate.Enabled = True
End Sub

Private Sub optNetClient_Click()
    cmbComBaud.Enabled = False
    cmbComID.Enabled = False
    cmbI2cClockRate.Enabled = False
End Sub

Private Sub optUart_Click()
    cmbComBaud.Enabled = True
    cmbComID.Enabled = True
    cmbI2cClockRate.Enabled = False
End Sub
