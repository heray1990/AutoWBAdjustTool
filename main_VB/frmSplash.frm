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
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
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
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Sample1"
      Top             =   1320
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
      Top             =   600
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Auto-Adjust ColorTemp System"
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
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   4365
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
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   " Copyright 2013-2015           Design by  TE            ATS INC."
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

Dim ss As Boolean

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
    ss = False
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
 
    sqlstring = "select * from CheckItem"
    Executesql (sqlstring)

    If rs.EOF = False Then
        rs.MoveFirst
        cmbModelName.Clear

        Do While Not rs.EOF
            cmbModelName.AddItem rs.Fields("Mark")
            rs.MoveNext
        Loop
    Else
        MsgBox "Read Data Error,Please Check Your Database!", vbOKOnly + vbInformation, "Warning!"
        End
    End If
    
    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""
   
    sqlstring = "select * from CommonTable where Mark='ATS'"
    Executesql (sqlstring)

    If rs.EOF = False Then
        strCurrentModelName = rs("CurrentModelName")
        strDataVersion = rs("DataVersion")
        SetTVCurrentComID = rs("ComID")
        SetData = rs("Date")
        SetDay = rs("Day")
    Else
        MsgBox "Read Data Error,Please Check Your Database!", vbOKOnly + vbInformation, "Warning!"
    End If

    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""

    cmbModelName = strCurrentModelName

    If SetData <> Day(Date) Then
        sqlstring = "select * from CommonTable where Mark='ATS'"
        Executesql (sqlstring)
        rs.Fields(4) = Day(Date)
        rs.Fields(5) = SetDay + 1
  
        rs.Update

        Set cn = Nothing
        Set rs = Nothing
        sqlstring = ""
    End If

    Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo ErrExit

    strCurrentModelName = cmbModelName
    sqlstring = ""

    sqlstring = "update CommonTable set CurrentModelName='" & strCurrentModelName & "' where Mark='ATS'"
    Executesql (sqlstring)
    
    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""

    sqlstring = "select * from CheckItem where Mark='" & strCurrentModelName & "'"
    Executesql (sqlstring)

    SetTVCurrentComBaud = rs("ComBaud")
    IsCa210Channel = rs("Channel")
    IsStepTime = rs("Delayms")
    IsWhitePtn = rs("WhitePattern")

    IsAdjCool_1 = rs("COOL_1")
    IsAdjCool_2 = rs("COOL_2")
    IsAdjNormal = rs("NORMAL")
    IsAdjWarm_1 = rs("WARM_1")
    IsAdjWarm_2 = rs("WARM_2")

    IsBarcodeLen = rs("SN_Len")
    IsSensorLight = rs("SensorL")
    IsSaveData = rs("SaveData")
    IsCheckColorTemp = rs("CheckColor")

    IsSendOffset = rs("SendOFF")
    IsAdjsutOffset = rs("AdjustOFF")

    IsCool_1ModeIndex = rs("Cool_1MI")
    IsCool_2ModeIndex = rs("Cool_2MI")
    IsNormalModeIndex = rs("NormalMI")
    IsWarm_1ModeIndex = rs("Warm_1MI")
    IsWarm_2ModeIndex = rs("Warm_2MI")

    Set rs = Nothing
    Set cn = Nothing
    sqlstring = ""

    Form1.Show
    Exit Sub

ErrExit:
    MsgBox ("The Licence Key is Wrong.")
End Sub

Private Function ATS() As Boolean
Dim path As String
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim oldkey As String
Dim i%
Dim key As Single


Dim fso As New FileSystemObject
Dim Hdid
Dim hardwareid As String
Set fso = CreateObject("scripting.filesystemobject")
Set Hdid = fso.GetDrive("C:")
hardwareid = Hex(Hdid.SerialNumber)
path = App.path

On Error GoTo SSS

kk:
ATS = False

Open ("C:\source.dll") For Input As #1
Input #1, b
Close #1
Open ("C:\sys.dat") For Input As #2
Input #2, c
Close #2


key = Val(b)



If key < 283155 And key > 282950 And c = "3" + hardwareid Then
  If Month(Date) > 4 And Month(Date) < 7 And Year(Date) = 2014 Then
    ATS = True
  Else
    a = Str$(key + 100)
  End If
  a = Str$(key + 1)
  Open ("C:\source.dll") For Output As #8
  Print #8, a
  Close #8
  Exit Function
Else
  If i > 1 Then
  MsgBox ("Please apply for a licensed APP.")
  Unload frmSplash
  End
  End If
  
  GoTo SSS
  End
  Exit Function
End If

SSS:


For i = 1 To 3

a = InputBox("" & vbNewLine & "     Please Input The Licence Key.", "LICENCE")
If a = "2829558" Then



a = Str$(Val(key) + Val(Left$(a, 6)))

Open ("C:\source.dll") For Output As #5
Print #5, a
Close #5



i = i + 1
GoTo kk
Exit Function
Else
  If a = "DIPHD@23456" Then MsgBox (hardwareid)
  MsgBox ("The Licence Key is Wrong.")
End If

Next i
  End
  Exit Function

End Function


