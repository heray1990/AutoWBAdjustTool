VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2730
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5610
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbModelName 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   600
      Left            =   1080
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Sample1"
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Version "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   720
      Width           =   825
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Auto White Balance System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   5355
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Please select your model:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Copyright 2013-2016    Design by ECHOM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   5535
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_DblClick()
    Unload Me
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()

On Error GoTo ErrExit
    Dim strProjectName As Variant
    
    cmbModelName.Clear
    
    For Each strProjectName In GetProjectList
        cmbModelName.AddItem strProjectName
    Next strProjectName
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

    cmbModelName.Text = GetCurProjectName
    
    Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub Form_Unload(Cancel As Integer)

'On Error GoTo ErrExit
    gstrCurProjName = cmbModelName.Text
    SetCurProjectName gstrCurProjName

    Form1.Show
    Exit Sub

'ErrExit:
    'MsgBox ("The Licence Key is Wrong.")
End Sub
