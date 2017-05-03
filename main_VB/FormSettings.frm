VERSION 5.00
Begin VB.Form FormSettings 
   Caption         =   "SpecData"
   ClientHeight    =   7350
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
   ScaleHeight     =   7350
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
      TabIndex        =   34
      Top             =   3720
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   35
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   41
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
            TabIndex        =   40
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
            TabIndex        =   39
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
         TabIndex        =   45
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
         TabIndex        =   44
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
      TabIndex        =   31
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
         TabIndex        =   32
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
         TabIndex        =   33
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
      TabIndex        =   25
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   29
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
         TabIndex        =   28
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
      TabIndex        =   22
      Top             =   840
      Width           =   2400
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
         TabIndex        =   30
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
         TabIndex        =   24
         Top             =   360
         Width           =   800
      End
      Begin VB.OptionButton optNetwork 
         Caption         =   "Network"
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
         TabIndex        =   23
         Top             =   360
         Width           =   1000
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
      TabIndex        =   10
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
         TabIndex        =   21
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
         TabIndex        =   17
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
         TabIndex        =   13
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
         TabIndex        =   11
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
         TabIndex        =   20
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
         TabIndex        =   18
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
         TabIndex        =   14
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
         TabIndex        =   12
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
      TabIndex        =   7
      Top             =   6480
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
         TabIndex        =   9
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
         TabIndex        =   8
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
      Height          =   2835
      Left            =   120
      TabIndex        =   6
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
         TabIndex        =   16
         Top             =   2450
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
         TabIndex        =   15
         Top             =   2100
         Value           =   1  'Checked
         Width           =   1900
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "COOL_2"
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
         Width           =   1900
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "COOL_1"
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
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   "NORMAL"
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
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "WARM_1"
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
         TabIndex        =   3
         Top             =   1400
         Value           =   1  'Checked
         Width           =   1900
      End
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         Caption         =   "WARM_2"
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
         TabIndex        =   4
         Top             =   1750
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
      TabIndex        =   5
      Top             =   6840
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
      TabIndex        =   19
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

    If gutdCommMode = modeUART Then
        optUart.Value = True
        optNetwork.Value = False
        optI2c.Value = False
        cmbComBaud.Enabled = True
        cmbComID.Enabled = True
        cmbI2cClockRate.Enabled = False
    ElseIf gutdCommMode = modeNetwork Then
        optUart.Value = False
        optNetwork.Value = True
        optI2c.Value = False
        cmbComBaud.Enabled = False
        cmbComID.Enabled = False
        cmbI2cClockRate.Enabled = False
    ElseIf gutdCommMode = modeI2c Then
        optUart.Value = False
        optNetwork.Value = False
        optI2c.Value = True
        cmbComBaud.Enabled = False
        cmbComID.Enabled = False
        cmbI2cClockRate.Enabled = True
    End If
    
    If gblnEnableCool2 Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If

    If gblnEnableCool1 Then
        Check2.Value = 1
    Else
        Check2.Value = 0
    End If

    If gblnEnableStandard Then
        Check3.Value = 1
    Else
        Check3.Value = 0
    End If

    If gblnEnableWarm1 Then
        Check4.Value = 1
    Else
        Check4.Value = 0
    End If

    If gblnEnableWarm2 Then
        Check5.Value = 1
    Else
        Check5.Value = 0
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
    Dim clsSaveConfigData As ProjectConfig
    
    Set clsSaveConfigData = New ProjectConfig
    
    If Check1.Value = 1 Then clsSaveConfigData.EnableCool2 = True
    If Check1.Value = 0 Then clsSaveConfigData.EnableCool2 = False
    If Check2.Value = 1 Then clsSaveConfigData.EnableCool1 = True
    If Check2.Value = 0 Then clsSaveConfigData.EnableCool1 = False
    If Check3.Value = 1 Then clsSaveConfigData.EnableNormal = True
    If Check3.Value = 0 Then clsSaveConfigData.EnableNormal = False
    If Check4.Value = 1 Then clsSaveConfigData.EnableWarm1 = True
    If Check4.Value = 0 Then clsSaveConfigData.EnableWarm1 = False
    If Check5.Value = 1 Then clsSaveConfigData.EnableWarm2 = True
    If Check5.Value = 0 Then clsSaveConfigData.EnableWarm2 = False
    If Check6.Value = 1 Then clsSaveConfigData.EnableChkColor = True
    If Check6.Value = 0 Then clsSaveConfigData.EnableChkColor = False
    If Check7.Value = 1 Then clsSaveConfigData.EnableAdjOffset = True
    If Check7.Value = 0 Then clsSaveConfigData.EnableAdjOffset = False
    
    clsSaveConfigData.LvSpec = val(txtLvSpec.Text)
    clsSaveConfigData.BarCodeLen = val(txtSNLen.Text)
    
    If optUart.Value = True Then
        clsSaveConfigData.CommMode = modeUART
    ElseIf optNetwork.Value = True Then
        clsSaveConfigData.CommMode = modeNetwork
    ElseIf optI2c.Value = True Then
        clsSaveConfigData.CommMode = modeI2c
    Else
        clsSaveConfigData.CommMode = modeUART
    End If

    clsSaveConfigData.ComBaud = cmbComBaud.Text
    clsSaveConfigData.ComID = val(Replace(cmbComID.Text, "COM", ""))
    clsSaveConfigData.I2cClockRate = val(Replace(cmbI2cClockRate.Text, "KHz", ""))
    clsSaveConfigData.ChannelNum = val(txtChannel.Text)
    clsSaveConfigData.DelayMS = val(txtDelay.Text)
    clsSaveConfigData.inputSource = cmbInputSource.Text
    clsSaveConfigData.VPGModel = cmbChromaModel.Text
    clsSaveConfigData.VPGTiming = txtChromaTiming.Text
    clsSaveConfigData.VPG100IRE = txt100IRE.Text
    clsSaveConfigData.VPG80IRE = txt80IRE.Text
    clsSaveConfigData.VPG20IRE = txt20IRE.Text
    
    clsSaveConfigData.SaveConfigData
    
    Set clsSaveConfigData = Nothing

    Unload Me
    
    FormMain.subInitInterface
    FormMain.Show
End Sub

Private Sub optI2c_Click()
    cmbComBaud.Enabled = False
    cmbComID.Enabled = False
    cmbI2cClockRate.Enabled = True
End Sub

Private Sub optNetwork_Click()
    cmbComBaud.Enabled = False
    cmbComID.Enabled = False
    cmbI2cClockRate.Enabled = False
End Sub

Private Sub optUart_Click()
    cmbComBaud.Enabled = True
    cmbComID.Enabled = True
    cmbI2cClockRate.Enabled = False
End Sub
