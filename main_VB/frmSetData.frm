VERSION 5.00
Begin VB.Form frmSetData 
   Caption         =   "SpecData"
   ClientHeight    =   4920
   ClientLeft      =   6435
   ClientTop       =   3210
   ClientWidth     =   4950
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
   ScaleHeight     =   4920
   ScaleWidth      =   4950
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
      Left            =   120
      TabIndex        =   23
      Top             =   4080
      Width           =   2295
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
         TabIndex        =   25
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
         TabIndex        =   24
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
      Left            =   2520
      TabIndex        =   10
      Top             =   1680
      Width           =   2400
      Begin VB.TextBox txtLvSpec 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   22
         Text            =   "280"
         Top             =   1350
         Width           =   1000
      End
      Begin VB.ComboBox cmbInputSource 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmSetData.frx":1DF72
         Left            =   1200
         List            =   "frmSetData.frx":1DF7F
         TabIndex        =   18
         Text            =   "HDMI1"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox txtSNLen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
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
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
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
      Begin VB.Label Label6 
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
         Index           =   1
         Left            =   200
         TabIndex        =   21
         Top             =   1380
         Width           =   900
      End
      Begin VB.Label Label5 
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
         Index           =   0
         Left            =   200
         TabIndex        =   19
         Top             =   1030
         Width           =   900
      End
      Begin VB.Label Label4 
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
         Index           =   0
         Left            =   200
         TabIndex        =   14
         Top             =   680
         Width           =   900
      End
      Begin VB.Label Label3 
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
         Index           =   0
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
      Left            =   2520
      TabIndex        =   7
      Top             =   840
      Width           =   2400
      Begin VB.TextBox txtChannel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
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
      Begin VB.Label Label2 
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
         Index           =   0
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
      Height          =   3200
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   2295
      Begin VB.CheckBox Check8 
         Alignment       =   1  'Right Justify
         Caption         =   "AdjustOffset"
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
         TabIndex        =   17
         Top             =   2800
         Value           =   1  'Checked
         Width           =   1800
      End
      Begin VB.CheckBox Check7 
         Alignment       =   1  'Right Justify
         Caption         =   "CheckColor"
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
         Width           =   1800
      End
      Begin VB.CheckBox Check6 
         Alignment       =   1  'Right Justify
         Caption         =   "Save Data"
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
         Width           =   1800
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
         Width           =   1800
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
         Width           =   1800
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
         Width           =   1800
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
         Width           =   1800
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
         Width           =   1800
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
      Left            =   3720
      TabIndex        =   5
      Top             =   4320
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
      TabIndex        =   20
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
    txtSNLen.Text = CStr(BarCodeLen)
    txtLvSpec.Text = CStr(maxBrightnessSpec)
    cmbInputSource.Text = setTVInputSource & CStr(setTVInputSourcePortNum)
    txtDelay.Text = delayTime

    If isUartMode Then
        optUart.Value = True
        optNetwork.Value = False
    Else
        optUart.Value = False
        optNetwork.Value = True
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

    If isSaveData Then
        Check6.Value = 1
    Else
        Check6.Value = 0
    End If

    If isCheckColorTemp Then
        Check7.Value = 1
    Else
        Check7.Value = 0
    End If

    If isAdjustOffset Then
        Check8.Value = 1
    Else
        Check8.Value = 0
    End If
    
End Sub

Private Sub Command1_Click()
    sqlstring = "select * from CheckItem where Mark='" & strCurrentModelName & "'"
    Executesql (sqlstring)

    If Check1.Value = 1 Then rs.Fields(1) = True
    If Check1.Value = 0 Then rs.Fields(1) = False
    If Check2.Value = 1 Then rs.Fields(2) = True
    If Check2.Value = 0 Then rs.Fields(2) = False
    If Check3.Value = 1 Then rs.Fields(3) = True
    If Check3.Value = 0 Then rs.Fields(3) = False
    If Check4.Value = 1 Then rs.Fields(4) = True
    If Check4.Value = 0 Then rs.Fields(4) = False
    If Check5.Value = 1 Then rs.Fields(5) = True
    If Check5.Value = 0 Then rs.Fields(5) = False
    If Check6.Value = 1 Then rs.Fields(6) = True
    If Check6.Value = 0 Then rs.Fields(6) = False
    If Check7.Value = 1 Then rs.Fields(7) = True
    If Check7.Value = 0 Then rs.Fields(7) = False
    If Check8.Value = 1 Then rs.Fields(8) = True
    If Check8.Value = 0 Then rs.Fields(8) = False
    
    rs.Fields(9) = Val(txtLvSpec.Text)
    rs.Fields(10) = Val(txtSNLen.Text)
    
    If optUart.Value = True Then
        rs.Fields(11) = "UART"
    ElseIf optNetwork.Value = True Then
        rs.Fields(11) = "Network"
    Else
        rs.Fields(11) = "UART"
    End If

    rs.Update

    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""

    sqlstring = "select * from CommonTable where Mark='ATS'"
    Executesql (sqlstring)

    rs.Fields(4) = Val(txtChannel.Text)
    rs.Fields(5) = Val(txtDelay.Text)
    rs.Fields(6) = cmbInputSource.Text
    
    rs.Update
    
    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""

    MsgBox "Save success!", vbOKOnly, "warning"
    Unload Me
    Unload Form1
End Sub

