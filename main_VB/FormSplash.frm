VERSION 5.00
Begin VB.Form FormSplash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2295
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FormSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   2895
      Picture         =   "FormSplash.frx":000C
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   4
      Top             =   120
      Width           =   780
   End
   Begin VB.PictureBox PictureBrand 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   360
      Picture         =   "FormSplash.frx":1E00
      ScaleHeight     =   750
      ScaleWidth      =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   2550
   End
   Begin VB.ComboBox cmbModelName 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Sample1"
      Top             =   1440
      Width           =   3300
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
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
      Left            =   2760
      TabIndex        =   2
      Top             =   1920
      Width           =   825
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Please select model:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   3255
   End
End
Attribute VB_Name = "FormSplash"
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

Private Sub cmbModelName_Click()
    SubUpdateBrand
End Sub

Private Sub Form_Load()
    On Error GoTo ErrExit
    Dim strProjectName As Variant
    
    cmbModelName.Clear
    
    For Each strProjectName In GetProjectList
        cmbModelName.AddItem strProjectName
    Next strProjectName
    
    lblVersion.Caption = TXTVersion & " " & App.Major & "." & App.Minor & "." & App.Revision
    Label1.Caption = TXTMessage1

    cmbModelName.Text = GetCurProjectName
    
    SubUpdateBrand
    
    Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo ErrExit
    gstrCurProjName = cmbModelName.Text
    SetCurProjectName gstrCurProjName

    FormMain.Show
    Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub SubUpdateBrand()
    Dim strBrand As String
    
    strBrand = Split(cmbModelName.Text, gstrDelimiterForProjName)(0)
    
    If UCase(strBrand) = "CAN" Then
        PictureBrand.Picture = LoadPicture(App.Path & "\Resources\CANTV.bmp")
    ElseIf UCase(strBrand) = "HAIER" Then
        PictureBrand.Picture = LoadPicture(App.Path & "\Resources\Haier.bmp")
    Else
        PictureBrand.Picture = LoadPicture(App.Path & "\Resources\Letv.bmp")
    End If
End Sub
