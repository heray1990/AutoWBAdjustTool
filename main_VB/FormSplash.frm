VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FormSplash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auto White Balance Tool"
   ClientHeight    =   1335
   ClientLeft      =   255
   ClientTop       =   1800
   ClientWidth     =   3390
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FormSplash.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CommandLoadXml 
      Caption         =   "加载配置文件"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   360
      Picture         =   "FormSplash.frx":000C
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   2
      Top             =   240
      Width           =   780
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Left            =   2280
      TabIndex        =   1
      Top             =   840
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
    Me.Caption = TXTTitle
    lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrExit
    FormMain.Show
    Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub
