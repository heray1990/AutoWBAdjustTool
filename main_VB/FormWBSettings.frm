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
   StartUpPosition =   3  'Windows Default
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
      Caption         =   "Ԥ��Offsetֵ"
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
      Caption         =   "Ԥ��Gainֵ"
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
      Caption         =   "������"
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
      Caption         =   "�������"
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
      Text            =   "Cool1"
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


Private Sub cmbColorT_Click()
    If cmbColorT.Text = "Cool1" Then
        txtSpecX.Text = Format(val(rConfigData.intSPECCool1x) / 10000, "0.####")
        txtSpecY.Text = Format(val(rConfigData.intSPECCool1y) / 10000, "0.####")
        txtTolX.Text = Format(val(rConfigData.intTOLCool1xt) / 10000, "0.####")
        txtTolY.Text = Format(val(rConfigData.intTOLCool1yt) / 10000, "0.####")
        txtChkX.Text = Format(val(rConfigData.intCHKCool1Cxt) / 10000, "0.####")
        txtChkY.Text = Format(val(rConfigData.intCHKCool1Cyt) / 10000, "0.####")
        txtPresetgainR.Text = CStr(rConfigData.intPRESETGANCool1R)
        txtPresetgainG.Text = CStr(rConfigData.intPRESETGANCool1G)
        txtPresetgainB.Text = CStr(rConfigData.intPRESETGANCool1B)
        txtPresetOffR.Text = CStr(rConfigData.intPRESETOFFCool1R)
        txtPresetOffG.Text = CStr(rConfigData.intPRESETOFFCool1G)
        txtPresetOffB.Text = CStr(rConfigData.intPRESETOFFCool1B)
        txtGainStepX.Text = Format(val(rConfigData.intMAGICVALGMin) / 10000, "0.####")
        txtGainStepY.Text = Format(val(rConfigData.intMAGICVALOMin) / 10000, "0.####")
        txtOffStepX.Text = Format(val(rConfigData.intMAGICVALGMax) / 10000, "0.####")
        txtOffStepY.Text = Format(val(rConfigData.intMAGICVALOMax) / 10000, "0.####")
    End If
    If cmbColorT.Text = "Normal" Then
        txtSpecX.Text = Format(val(rConfigData.intSPECNormalx) / 10000, "0.####")
        txtSpecY.Text = Format(val(rConfigData.intSPECNormaly) / 10000, "0.####")
        txtTolX.Text = Format(val(rConfigData.intTOLNormalxt) / 10000, "0.####")
        txtTolY.Text = Format(val(rConfigData.intTOLNormalyt) / 10000, "0.####")
        txtChkX.Text = Format(val(rConfigData.intCHKNormalCxt) / 10000, "0.####")
        txtChkY.Text = Format(val(rConfigData.intCHKNormalCyt) / 10000, "0.####")
        txtPresetgainR.Text = CStr(rConfigData.intPRESETGANNormalR)
        txtPresetgainG.Text = CStr(rConfigData.intPRESETGANNormalG)
        txtPresetgainB.Text = CStr(rConfigData.intPRESETGANNormalB)
        txtPresetOffR.Text = CStr(rConfigData.intPRESETOFFNormalR)
        txtPresetOffG.Text = CStr(rConfigData.intPRESETOFFNormalG)
        txtPresetOffB.Text = CStr(rConfigData.intPRESETOFFNormalB)
        txtGainStepX.Text = Format(val(rConfigData.intMAGICVALGMin) / 10000, "0.####")
        txtGainStepY.Text = Format(val(rConfigData.intMAGICVALOMin) / 10000, "0.####")
        txtOffStepX.Text = Format(val(rConfigData.intMAGICVALGMax) / 10000, "0.####")
        txtOffStepY.Text = Format(val(rConfigData.intMAGICVALOMax) / 10000, "0.####")
    End If
    If cmbColorT.Text = "Warm1" Then
        txtSpecX.Text = Format(val(rConfigData.intSPECWarm1x) / 10000, "0.####")
        txtSpecY.Text = Format(val(rConfigData.intSPECWarm1y) / 10000, "0.####")
        txtTolX.Text = Format(val(rConfigData.intTOLWarm1xt) / 10000, "0.####")
        txtTolY.Text = Format(val(rConfigData.intTOLWarm1yt) / 10000, "0.####")
        txtChkX.Text = Format(val(rConfigData.intCHKWarm1Cxt) / 10000, "0.####")
        txtChkY.Text = Format(val(rConfigData.intCHKWarm1Cyt) / 10000, "0.####")
        txtPresetgainR.Text = CStr(rConfigData.intPRESETGANWarm1R)
        txtPresetgainG.Text = CStr(rConfigData.intPRESETGANWarm1G)
        txtPresetgainB.Text = CStr(rConfigData.intPRESETGANWarm1B)
        txtPresetOffR.Text = CStr(rConfigData.intPRESETOFFWarm1R)
        txtPresetOffG.Text = CStr(rConfigData.intPRESETOFFWarm1G)
        txtPresetOffB.Text = CStr(rConfigData.intPRESETOFFWarm1B)
        txtGainStepX.Text = Format(val(rConfigData.intMAGICVALGMin) / 10000, "0.####")
        txtGainStepY.Text = Format(val(rConfigData.intMAGICVALOMin) / 10000, "0.####")
        txtOffStepX.Text = Format(val(rConfigData.intMAGICVALGMax) / 10000, "0.####")
        txtOffStepY.Text = Format(val(rConfigData.intMAGICVALOMax) / 10000, "0.####")
        End If
End Sub

Private Sub Form_Load()
    txtSpecX.Text = Format(val(rConfigData.intSPECCool1x) / 10000, "0.####")
    txtSpecY.Text = Format(val(rConfigData.intSPECCool1y) / 10000, "0.####")
    txtTolX.Text = Format(val(rConfigData.intTOLCool1xt) / 10000, "0.####")
    txtTolY.Text = Format(val(rConfigData.intTOLCool1yt) / 10000, "0.####")
    txtChkX.Text = Format(val(rConfigData.intCHKCool1Cxt) / 10000, "0.####")
    txtChkY.Text = Format(val(rConfigData.intCHKCool1Cyt) / 10000, "0.####")
    txtPresetgainR.Text = CStr(rConfigData.intPRESETGANCool1R)
    txtPresetgainG.Text = CStr(rConfigData.intPRESETGANCool1G)
    txtPresetgainB.Text = CStr(rConfigData.intPRESETGANCool1B)
    txtPresetOffR.Text = CStr(rConfigData.intPRESETOFFCool1R)
    txtPresetOffG.Text = CStr(rConfigData.intPRESETOFFCool1G)
    txtPresetOffB.Text = CStr(rConfigData.intPRESETOFFCool1B)
    txtGainStepX.Text = Format(val(rConfigData.intMAGICVALGMin) / 10000, "0.####")
    txtGainStepY.Text = Format(val(rConfigData.intMAGICVALOMin) / 10000, "0.####")
    txtOffStepX.Text = Format(val(rConfigData.intMAGICVALGMax) / 10000, "0.####")
    txtOffStepY.Text = Format(val(rConfigData.intMAGICVALOMax) / 10000, "0.####")
    cmbColorT.AddItem "Cool1"
    cmbColorT.AddItem "Normal"
    cmbColorT.AddItem "Warm1"
    Label1.Caption = gstrCurProjName
    
    
End Sub

Private Sub Command1_Click()
    If cmbColorT.Text = "Cool1" Then
        rConfigData.intSPECCool1x = val(txtSpecX.Text) * 10000
        rConfigData.intSPECCool1y = val(txtSpecY.Text) * 10000
        rConfigData.intTOLCool1xt = val(txtTolX.Text) * 10000
        rConfigData.intTOLCool1yt = val(txtTolY.Text) * 10000
        rConfigData.intCHKCool1Cxt = val(txtChkX.Text) * 10000
        rConfigData.intCHKCool1Cyt = val(txtChkY.Text) * 10000
        rConfigData.intPRESETGANCool1R = val(txtPresetgainR.Text)
        rConfigData.intPRESETGANCool1G = val(txtPresetgainG.Text)
        rConfigData.intPRESETGANCool1B = val(txtPresetgainB.Text)
        rConfigData.intPRESETOFFCool1R = val(txtPresetOffR.Text)
        rConfigData.intPRESETOFFCool1G = val(txtPresetOffG.Text)
        rConfigData.intPRESETOFFCool1B = val(txtPresetOffB.Text)
        rConfigData.intMAGICVALGMin = val(txtGainStepX.Text) * 10000
        rConfigData.intMAGICVALOMin = val(txtGainStepY.Text) * 10000
        rConfigData.intMAGICVALGMax = val(txtOffStepX.Text) * 10000
        rConfigData.intMAGICVALOMax = val(txtOffStepY.Text) * 10000
    End If
    If cmbColorT.Text = "Normal" Then
        rConfigData.intSPECNormalx = val(txtSpecX.Text) * 10000
        rConfigData.intSPECNormaly = val(txtSpecY.Text) * 10000
        rConfigData.intTOLNormalxt = val(txtTolX.Text) * 10000
        rConfigData.intTOLNormalyt = val(txtTolY.Text) * 10000
        rConfigData.intCHKNormalCxt = val(txtChkX.Text) * 10000
        rConfigData.intCHKNormalCyt = val(txtChkY.Text) * 10000
        rConfigData.intPRESETGANNormalR = val(txtPresetgainR.Text)
        rConfigData.intPRESETGANNormalG = val(txtPresetgainG.Text)
        rConfigData.intPRESETGANNormalB = val(txtPresetgainB.Text)
        rConfigData.intPRESETOFFNormalR = val(txtPresetOffR.Text)
        rConfigData.intPRESETOFFNormalG = val(txtPresetOffG.Text)
        rConfigData.intPRESETOFFNormalB = val(txtPresetOffB.Text)
        rConfigData.intMAGICVALGMin = val(txtGainStepX.Text) * 10000
        rConfigData.intMAGICVALOMin = val(txtGainStepY.Text) * 10000
        rConfigData.intMAGICVALGMax = val(txtOffStepX.Text) * 10000
        rConfigData.intMAGICVALOMax = val(txtOffStepY.Text) * 10000
    End If
    If cmbColorT.Text = "Warm1" Then
        rConfigData.intSPECWarm1x = val(txtSpecX.Text) * 10000
        rConfigData.intSPECWarm1y = val(txtSpecY.Text) * 10000
        rConfigData.intTOLWarm1xt = val(txtTolX.Text) * 10000
        rConfigData.intTOLWarm1yt = val(txtTolY.Text) * 10000
        rConfigData.intCHKWarm1Cxt = val(txtChkX.Text) * 10000
        rConfigData.intCHKWarm1Cyt = val(txtChkY.Text) * 10000
        rConfigData.intPRESETGANWarm1R = val(txtPresetgainR.Text)
        rConfigData.intPRESETGANWarm1G = val(txtPresetgainG.Text)
        rConfigData.intPRESETGANWarm1B = val(txtPresetgainB.Text)
        rConfigData.intPRESETOFFWarm1R = val(txtPresetOffR.Text)
        rConfigData.intPRESETOFFWarm1G = val(txtPresetOffG.Text)
        rConfigData.intPRESETOFFWarm1B = val(txtPresetOffB.Text)
        rConfigData.intMAGICVALGMin = val(txtGainStepX.Text) * 10000
        rConfigData.intMAGICVALOMin = val(txtGainStepY.Text) * 10000
        rConfigData.intMAGICVALGMax = val(txtOffStepX.Text) * 10000
        rConfigData.intMAGICVALOMax = val(txtOffStepY.Text) * 10000
    End If
    
    SaveConfigData1
    

    Unload Me
    
    FormMain.subInitInterface
    FormMain.Show
End Sub


