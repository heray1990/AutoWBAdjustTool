VERSION 5.00
Begin VB.Form FormWBSettings 
   Caption         =   "White Balance Settings"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5325
   Icon            =   "FormWBSettings.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
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
      Height          =   495
      Left            =   3840
      TabIndex        =   41
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Frame Frame6 
      Caption         =   "Majic Value"
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
      Left            =   2760
      TabIndex        =   8
      Top             =   5400
      Width           =   2415
      Begin VB.TextBox txtGainStepX 
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
         Left            =   1320
         TabIndex        =   40
         Text            =   "0.0015"
         Top             =   240
         Width           =   1000
      End
      Begin VB.TextBox txtGainStepY 
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
         Left            =   1320
         TabIndex        =   39
         Text            =   "0.002"
         Top             =   600
         Width           =   1000
      End
      Begin VB.TextBox txtOffStepX 
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
         Left            =   1320
         TabIndex        =   38
         Text            =   "0.0015"
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox txtOffStepY 
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
         Left            =   1320
         TabIndex        =   37
         Text            =   "0.002"
         Top             =   1320
         Width           =   1000
      End
      Begin VB.Label lbOffStepY 
         Alignment       =   1  'Right Justify
         Caption         =   "y:"
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
         TabIndex        =   24
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lbOffStepX 
         Caption         =   "Offset Step x:"
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
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lbGainStepY 
         Alignment       =   1  'Right Justify
         Caption         =   "y:"
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
         TabIndex        =   22
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbGainStepX 
         Alignment       =   1  'Right Justify
         Caption         =   "Gain Step x:"
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
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "‘§…ËOffset÷µ"
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
      Left            =   2760
      TabIndex        =   7
      Top             =   3360
      Width           =   2415
      Begin VB.TextBox txtPresetOffR 
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
         TabIndex        =   36
         Text            =   "128"
         Top             =   360
         Width           =   1000
      End
      Begin VB.TextBox txtPresetOffG 
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
         TabIndex        =   35
         Text            =   "128"
         Top             =   840
         Width           =   1000
      End
      Begin VB.TextBox txtPresetOffB 
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
         TabIndex        =   34
         Text            =   "128"
         Top             =   1320
         Width           =   1000
      End
      Begin VB.Label lbPresetOffB 
         Alignment       =   1  'Right Justify
         Caption         =   "B:"
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
         Left            =   840
         TabIndex        =   20
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lbPresetOffG 
         Alignment       =   1  'Right Justify
         Caption         =   "G:"
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
         Left            =   840
         TabIndex        =   19
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lbPresetOffR 
         Alignment       =   1  'Right Justify
         Caption         =   "R:"
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
         Left            =   840
         TabIndex        =   18
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "‘§…ËGain÷µ"
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
      Left            =   2760
      TabIndex        =   6
      Top             =   1320
      Width           =   2415
      Begin VB.TextBox txtPresetgainR 
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
         TabIndex        =   33
         Text            =   "128"
         Top             =   360
         Width           =   1000
      End
      Begin VB.TextBox txtPresetgainG 
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
         TabIndex        =   32
         Text            =   "128"
         Top             =   840
         Width           =   1000
      End
      Begin VB.TextBox txtPresetgainB 
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
         TabIndex        =   31
         Text            =   "128"
         Top             =   1320
         Width           =   1000
      End
      Begin VB.Label lbPresetgainB 
         Alignment       =   1  'Right Justify
         Caption         =   "B:"
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
         Left            =   840
         TabIndex        =   17
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lbPresetgainG 
         Alignment       =   1  'Right Justify
         Caption         =   "G:"
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
         Left            =   840
         TabIndex        =   16
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lbPresetgainR 
         Alignment       =   1  'Right Justify
         Caption         =   "R:"
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
         Left            =   840
         TabIndex        =   15
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "ºÏ≤ÈŒÛ≤Ó"
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
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   2415
      Begin VB.TextBox txtChkX 
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
         Left            =   1080
         TabIndex        =   30
         Text            =   "0.003"
         Top             =   600
         Width           =   1000
      End
      Begin VB.TextBox txtChkY 
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
         Left            =   1080
         TabIndex        =   29
         Text            =   "0.003"
         Top             =   1080
         Width           =   1000
      End
      Begin VB.Label lbChkX 
         Alignment       =   1  'Right Justify
         Caption         =   "x:"
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
         Left            =   720
         TabIndex        =   14
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lbChkY 
         Alignment       =   1  'Right Justify
         Caption         =   "y:"
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
         Left            =   720
         TabIndex        =   13
         Top             =   1080
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "≤‚ ‘ŒÛ≤Ó"
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
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   2415
      Begin VB.TextBox txtTolX 
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
         Left            =   1080
         TabIndex        =   28
         Text            =   "0.003"
         Top             =   600
         Width           =   1000
      End
      Begin VB.TextBox txtTolY 
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
         Left            =   1080
         TabIndex        =   27
         Text            =   "0.003"
         Top             =   1080
         Width           =   1000
      End
      Begin VB.Label lbTolY 
         Alignment       =   1  'Right Justify
         Caption         =   "y:"
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
         Left            =   720
         TabIndex        =   12
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lbTolX 
         Alignment       =   1  'Right Justify
         Caption         =   "x:"
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
         Left            =   720
         TabIndex        =   11
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Spec"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
      Begin VB.TextBox txtSpecY 
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
         Left            =   1080
         TabIndex        =   26
         Text            =   "0.278"
         Top             =   1080
         Width           =   1000
      End
      Begin VB.TextBox txtSpecX 
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
         Left            =   1080
         TabIndex        =   25
         Text            =   "0.272"
         Top             =   600
         Width           =   1000
      End
      Begin VB.Label lbSpecY 
         Alignment       =   1  'Right Justify
         Caption         =   "y:"
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
         Left            =   720
         TabIndex        =   10
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lbSpecX 
         Alignment       =   1  'Right Justify
         Caption         =   "x:"
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
         Left            =   720
         TabIndex        =   9
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.ComboBox cmbColorT 
      Height          =   315
      ItemData        =   "FormWBSettings.frx":1DF72
      Left            =   1800
      List            =   "FormWBSettings.frx":1DF74
      TabIndex        =   2
      Text            =   "COOL1"
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lbColorT 
      Caption         =   "Color Temperature:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1695
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
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "FormWBSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    txtSpecX.Text = Format(val(gudtSpecData.intSPECCool1x) / 10000, "0.####")
    txtSpecY.Text = Format(val(gudtSpecData.intSPECCool1y) / 10000, "0.####")
    txtTolX.Text = Format(val(gudtSpecData.intTOLCool1xt) / 10000, "0.####")
    txtTolY.Text = Format(val(gudtSpecData.intTOLCool1yt) / 10000, "0.####")
    txtChkX.Text = Format(val(gudtSpecData.intCHKCool1Cxt) / 10000, "0.####")
    txtChkY.Text = Format(val(gudtSpecData.intCHKCool1Cyt) / 10000, "0.####")
    txtPresetgainR.Text = CStr(gudtSpecData.intPRESETGANCool1R)
    txtPresetgainG.Text = CStr(gudtSpecData.intPRESETGANCool1G)
    txtPresetgainB.Text = CStr(gudtSpecData.intPRESETGANCool1B)
    txtPresetOffR.Text = CStr(gudtSpecData.intPRESETOFFCool1R)
    txtPresetOffG.Text = CStr(gudtSpecData.intPRESETOFFCool1G)
    txtPresetOffB.Text = CStr(gudtSpecData.intPRESETOFFCool1B)
    txtGainStepX.Text = Format(val(gudtSpecData.intMAGICVALGMin) / 10000, "0.####")
    txtGainStepY.Text = Format(val(gudtSpecData.intMAGICVALOMin) / 10000, "0.####")
    txtOffStepX.Text = Format(val(gudtSpecData.intMAGICVALGMax) / 10000, "0.####")
    txtOffStepY.Text = Format(val(gudtSpecData.intMAGICVALOMax) / 10000, "0.####")
    cmbColorT.AddItem COLORTEMP_COOL1
    cmbColorT.AddItem COLORTEMP_STANDARD
    cmbColorT.AddItem COLORTEMP_WARM1
    Label1.Caption = gstrCurProjName
End Sub

Private Sub cmbColorT_Click()
    If cmbColorT.Text = COLORTEMP_COOL1 Then
        txtSpecX.Text = Format(val(gudtSpecData.intSPECCool1x) / 10000, "0.####")
        txtSpecY.Text = Format(val(gudtSpecData.intSPECCool1y) / 10000, "0.####")
        txtTolX.Text = Format(val(gudtSpecData.intTOLCool1xt) / 10000, "0.####")
        txtTolY.Text = Format(val(gudtSpecData.intTOLCool1yt) / 10000, "0.####")
        txtChkX.Text = Format(val(gudtSpecData.intCHKCool1Cxt) / 10000, "0.####")
        txtChkY.Text = Format(val(gudtSpecData.intCHKCool1Cyt) / 10000, "0.####")
        txtPresetgainR.Text = CStr(gudtSpecData.intPRESETGANCool1R)
        txtPresetgainG.Text = CStr(gudtSpecData.intPRESETGANCool1G)
        txtPresetgainB.Text = CStr(gudtSpecData.intPRESETGANCool1B)
        txtPresetOffR.Text = CStr(gudtSpecData.intPRESETOFFCool1R)
        txtPresetOffG.Text = CStr(gudtSpecData.intPRESETOFFCool1G)
        txtPresetOffB.Text = CStr(gudtSpecData.intPRESETOFFCool1B)
    ElseIf cmbColorT.Text = COLORTEMP_STANDARD Then
        txtSpecX.Text = Format(val(gudtSpecData.intSPECNormalx) / 10000, "0.####")
        txtSpecY.Text = Format(val(gudtSpecData.intSPECNormaly) / 10000, "0.####")
        txtTolX.Text = Format(val(gudtSpecData.intTOLNormalxt) / 10000, "0.####")
        txtTolY.Text = Format(val(gudtSpecData.intTOLNormalyt) / 10000, "0.####")
        txtChkX.Text = Format(val(gudtSpecData.intCHKNormalCxt) / 10000, "0.####")
        txtChkY.Text = Format(val(gudtSpecData.intCHKNormalCyt) / 10000, "0.####")
        txtPresetgainR.Text = CStr(gudtSpecData.intPRESETGANNormalR)
        txtPresetgainG.Text = CStr(gudtSpecData.intPRESETGANNormalG)
        txtPresetgainB.Text = CStr(gudtSpecData.intPRESETGANNormalB)
        txtPresetOffR.Text = CStr(gudtSpecData.intPRESETOFFNormalR)
        txtPresetOffG.Text = CStr(gudtSpecData.intPRESETOFFNormalG)
        txtPresetOffB.Text = CStr(gudtSpecData.intPRESETOFFNormalB)
    ElseIf cmbColorT.Text = COLORTEMP_WARM1 Then
        txtSpecX.Text = Format(val(gudtSpecData.intSPECWarm1x) / 10000, "0.####")
        txtSpecY.Text = Format(val(gudtSpecData.intSPECWarm1y) / 10000, "0.####")
        txtTolX.Text = Format(val(gudtSpecData.intTOLWarm1xt) / 10000, "0.####")
        txtTolY.Text = Format(val(gudtSpecData.intTOLWarm1yt) / 10000, "0.####")
        txtChkX.Text = Format(val(gudtSpecData.intCHKWarm1Cxt) / 10000, "0.####")
        txtChkY.Text = Format(val(gudtSpecData.intCHKWarm1Cyt) / 10000, "0.####")
        txtPresetgainR.Text = CStr(gudtSpecData.intPRESETGANWarm1R)
        txtPresetgainG.Text = CStr(gudtSpecData.intPRESETGANWarm1G)
        txtPresetgainB.Text = CStr(gudtSpecData.intPRESETGANWarm1B)
        txtPresetOffR.Text = CStr(gudtSpecData.intPRESETOFFWarm1R)
        txtPresetOffG.Text = CStr(gudtSpecData.intPRESETOFFWarm1G)
        txtPresetOffB.Text = CStr(gudtSpecData.intPRESETOFFWarm1B)
    End If

    txtGainStepX.Text = Format(val(gudtSpecData.intMAGICVALGMin) / 10000, "0.####")
    txtGainStepY.Text = Format(val(gudtSpecData.intMAGICVALOMin) / 10000, "0.####")
    txtOffStepX.Text = Format(val(gudtSpecData.intMAGICVALGMax) / 10000, "0.####")
    txtOffStepY.Text = Format(val(gudtSpecData.intMAGICVALOMax) / 10000, "0.####")
End Sub

Private Sub Command1_Click()
    If cmbColorT.Text = COLORTEMP_COOL1 Then
        gudtSpecData.intSPECCool1x = val(txtSpecX.Text) * 10000
        gudtSpecData.intSPECCool1y = val(txtSpecY.Text) * 10000
        gudtSpecData.intTOLCool1xt = val(txtTolX.Text) * 10000
        gudtSpecData.intTOLCool1yt = val(txtTolY.Text) * 10000
        gudtSpecData.intCHKCool1Cxt = val(txtChkX.Text) * 10000
        gudtSpecData.intCHKCool1Cyt = val(txtChkY.Text) * 10000
        gudtSpecData.intPRESETGANCool1R = val(txtPresetgainR.Text)
        gudtSpecData.intPRESETGANCool1G = val(txtPresetgainG.Text)
        gudtSpecData.intPRESETGANCool1B = val(txtPresetgainB.Text)
        gudtSpecData.intPRESETOFFCool1R = val(txtPresetOffR.Text)
        gudtSpecData.intPRESETOFFCool1G = val(txtPresetOffG.Text)
        gudtSpecData.intPRESETOFFCool1B = val(txtPresetOffB.Text)
    ElseIf cmbColorT.Text = COLORTEMP_STANDARD Then
        gudtSpecData.intSPECNormalx = val(txtSpecX.Text) * 10000
        gudtSpecData.intSPECNormaly = val(txtSpecY.Text) * 10000
        gudtSpecData.intTOLNormalxt = val(txtTolX.Text) * 10000
        gudtSpecData.intTOLNormalyt = val(txtTolY.Text) * 10000
        gudtSpecData.intCHKNormalCxt = val(txtChkX.Text) * 10000
        gudtSpecData.intCHKNormalCyt = val(txtChkY.Text) * 10000
        gudtSpecData.intPRESETGANNormalR = val(txtPresetgainR.Text)
        gudtSpecData.intPRESETGANNormalG = val(txtPresetgainG.Text)
        gudtSpecData.intPRESETGANNormalB = val(txtPresetgainB.Text)
        gudtSpecData.intPRESETOFFNormalR = val(txtPresetOffR.Text)
        gudtSpecData.intPRESETOFFNormalG = val(txtPresetOffG.Text)
        gudtSpecData.intPRESETOFFNormalB = val(txtPresetOffB.Text)
    ElseIf cmbColorT.Text = COLORTEMP_WARM1 Then
        gudtSpecData.intSPECWarm1x = val(txtSpecX.Text) * 10000
        gudtSpecData.intSPECWarm1y = val(txtSpecY.Text) * 10000
        gudtSpecData.intTOLWarm1xt = val(txtTolX.Text) * 10000
        gudtSpecData.intTOLWarm1yt = val(txtTolY.Text) * 10000
        gudtSpecData.intCHKWarm1Cxt = val(txtChkX.Text) * 10000
        gudtSpecData.intCHKWarm1Cyt = val(txtChkY.Text) * 10000
        gudtSpecData.intPRESETGANWarm1R = val(txtPresetgainR.Text)
        gudtSpecData.intPRESETGANWarm1G = val(txtPresetgainG.Text)
        gudtSpecData.intPRESETGANWarm1B = val(txtPresetgainB.Text)
        gudtSpecData.intPRESETOFFWarm1R = val(txtPresetOffR.Text)
        gudtSpecData.intPRESETOFFWarm1G = val(txtPresetOffG.Text)
        gudtSpecData.intPRESETOFFWarm1B = val(txtPresetOffB.Text)
    End If
    
    gudtSpecData.intMAGICVALGMin = val(txtGainStepX.Text) * 10000
    gudtSpecData.intMAGICVALOMin = val(txtGainStepY.Text) * 10000
    gudtSpecData.intMAGICVALGMax = val(txtOffStepX.Text) * 10000
    gudtSpecData.intMAGICVALOMax = val(txtOffStepY.Text) * 10000
    
    Call SaveSpecData(cmbColorT.Text)

    Unload Me
    
    FormMain.SubInit
    FormMain.Show
End Sub


