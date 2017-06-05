VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CommandLoadXml 
      Caption         =   "加载配置文件"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   120
      Width           =   2550
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
      TabIndex        =   0
      Top             =   1920
      Width           =   825
   End
End
Attribute VB_Name = "FormSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CommandLoadXml_Click()
    ' CancelError is True.
    On Error GoTo ErrHandler
    ' Set filters.
    CommonDialog1.Filter = "All Files (*.*)|*.*|Xml Files (*.xml)|*.xml"
    ' Specify default filter.
    CommonDialog1.FilterIndex = 2

    ' Display the Open dialog box.
    CommonDialog1.ShowOpen
    gstrXmlPath = CommonDialog1.FileName
    
    Unload Me
    Exit Sub

ErrHandler:
    ' User pressed Cancel button.
    MsgBox "请加载 XML 文件，否则无法运行软件", vbExclamation, "加载配置文件"
    Exit Sub
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrExit
    FormMain.Show
    Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub
