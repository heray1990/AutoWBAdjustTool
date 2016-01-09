VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "SetComport"
   ClientHeight    =   1695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3990
   Icon            =   "frmComPort.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1695
   ScaleWidth      =   3990
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame Frame1 
      Caption         =   "ComSet"
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton cmdSet 
         Caption         =   "Set"
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Caption         =   "TV"
         ForeColor       =   &H000000C0&
         Height          =   1095
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2175
         Begin VB.ComboBox cmbTbaud 
            Height          =   300
            Left            =   1080
            TabIndex        =   3
            Text            =   "9600"
            Top             =   660
            Width           =   975
         End
         Begin VB.ComboBox cmbTcomID 
            Height          =   300
            ItemData        =   "frmComPort.frx":1DF72
            Left            =   1080
            List            =   "frmComPort.frx":1DF74
            TabIndex        =   2
            Text            =   "COM1"
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "ComBaud:"
            Height          =   255
            Index           =   0
            Left            =   200
            TabIndex        =   5
            Top             =   700
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "ComID:"
            Height          =   255
            Index           =   0
            Left            =   200
            TabIndex        =   4
            Top             =   360
            Width           =   900
         End
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
    Unload Me
    Form1.ZOrder (0)
End Sub

Private Sub cmdSet_Click()
    On Error GoTo ErrExit

    If Len(Trim$(cmbTcomID.Text)) = 5 Then
        SetTVCurrentComID = Val(Right(Trim$(cmbTcomID.Text), 2))
    ElseIf Len(Trim$(cmbTcomID.Text)) = 4 Then
        SetTVCurrentComID = Val(Right(Trim$(cmbTcomID.Text), 1))
    Else
        SetTVCurrentComID = 1
    End If
 
    SetTVCurrentComBaud = Val(cmbTbaud)

    If Form1.MSComm1.PortOpen = True Then
        Form1.MSComm1.PortOpen = False
    End If

    With Form1
        .MSComm1.CommPort = SetTVCurrentComID
        .MSComm1.Settings = SetTVCurrentComBaud & ",N,8,1"
        .MSComm1.PortOpen = True
    End With

    Unload Me
    Form1.ZOrder (0)
    Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub Form_Load()
On Error GoTo ErrExit

    cmbTcomID.Text = "COM" & SetTVCurrentComID
    cmbTbaud.Text = SetTVCurrentComBaud

    For i = 1 To 20
        cmbTcomID.AddItem "COM" & i
    Next i

    cmbTbaud.AddItem "9600"
    cmbTbaud.AddItem "19200"
    cmbTbaud.AddItem "38400"
    cmbTbaud.AddItem "57600"
    cmbTbaud.AddItem "115200"

    Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub



