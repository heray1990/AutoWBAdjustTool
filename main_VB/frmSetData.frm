VERSION 5.00
Begin VB.Form frmSetData 
   Caption         =   "SpecData"
   ClientHeight    =   5655
   ClientLeft      =   6435
   ClientTop       =   3210
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetData.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   5655
   ScaleWidth      =   5055
   Begin VB.Frame Frame1 
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
      Top             =   2520
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
         ItemData        =   "frmSetData.frx":1DF72
         Left            =   1200
         List            =   "frmSetData.frx":1DF74
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
         Left            =   1200
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
   Begin VB.Frame Frame3 
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
      Height          =   795
      Left            =   2550
      TabIndex        =   22
      Top             =   1680
      Width           =   2400
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
   Begin VB.Frame Frame5 
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
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   3720
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
         ItemData        =   "frmSetData.frx":1DF76
         Left            =   1200
         List            =   "frmSetData.frx":1DF83
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
   Begin VB.Frame Frame5 
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
      Index           =   0
      Left            =   2550
      TabIndex        =   7
      Top             =   840
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
   Begin VB.Frame Selection 
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
      Top             =   5040
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
      Width           =   4695
   End
End
Attribute VB_Name = "frmSetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    Label1.Caption = strCurrentModelName
    
    txtChannel.Text = CStr(Ca210ChannelNO)
    txtSNLen.Text = CStr(barCodeLen)
    txtLvSpec.Text = CStr(maxBrightnessSpec)
    cmbInputSource.Text = setTVInputSource & CStr(setTVInputSourcePortNum)
    txtDelay.Text = delayTime

    cmbComBaud.Text = CStr(setTVCurrentComBaud)
    cmbComID.Text = "COM" & CStr(setTVCurrentComID)
    For i = 1 To 20
        cmbComID.AddItem "COM" & i
    Next i

    cmbComBaud.AddItem "9600"
    cmbComBaud.AddItem "19200"
    cmbComBaud.AddItem "38400"
    cmbComBaud.AddItem "57600"
    cmbComBaud.AddItem "115200"

    If isUartMode Then
        optUart.Value = True
        optNetwork.Value = False
        cmbComBaud.Enabled = True
        cmbComID.Enabled = True
    Else
        optUart.Value = False
        optNetwork.Value = True
        cmbComBaud.Enabled = False
        cmbComID.Enabled = False
    End If
    
    If isAdjustCool2 Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If

    If isAdjustCool1 Then
        Check2.Value = 1
    Else
        Check2.Value = 0
    End If

    If isAdjustNormal Then
        Check3.Value = 1
    Else
        Check3.Value = 0
    End If

    If isAdjustWarm1 Then
        Check4.Value = 1
    Else
        Check4.Value = 0
    End If

    If isAdjustWarm2 Then
        Check5.Value = 1
    Else
        Check5.Value = 0
    End If

    If isCheckColorTemp Then
        Check6.Value = 1
    Else
        Check6.Value = 0
    End If

    If isAdjustOffset Then
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
    
    clsSaveConfigData.LvSpec = Val(txtLvSpec.Text)
    clsSaveConfigData.barCodeLen = Val(txtSNLen.Text)
    
    If optUart.Value = True Then
        clsSaveConfigData.CommMode = modeUART
    ElseIf optNetwork.Value = True Then
        clsSaveConfigData.CommMode = modeNetwork
    Else
        clsSaveConfigData.CommMode = modeUART
    End If

    clsSaveConfigData.ComBaud = cmbComBaud.Text
    clsSaveConfigData.ComID = Val(Replace(cmbComID.Text, "COM", ""))
    clsSaveConfigData.ChannelNum = Val(txtChannel.Text)
    clsSaveConfigData.DelayMS = Val(txtDelay.Text)
    clsSaveConfigData.inputSource = cmbInputSource.Text
    
    clsSaveConfigData.SaveConfigData
    
    Set clsSaveConfigData = Nothing

    MsgBox "Save success!", vbOKOnly, "warning"
    Unload Me
    Unload Form1
End Sub

Private Sub optNetwork_Click()
    cmbComBaud.Enabled = False
    cmbComID.Enabled = False
End Sub

Private Sub optUart_Click()
    cmbComBaud.Enabled = True
    cmbComID.Enabled = True
End Sub
