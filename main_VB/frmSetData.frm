VERSION 5.00
Begin VB.Form frmSetData 
   Caption         =   "SpecData"
   ClientHeight    =   4020
   ClientLeft      =   6435
   ClientTop       =   3210
   ClientWidth     =   10950
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   4020
   ScaleWidth      =   10950
   Begin VB.Frame Frame1 
      Caption         =   "CommSetting"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   0
      TabIndex        =   28
      Top             =   600
      Width           =   3135
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   34
         Text            =   "103"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Text            =   "1"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   1
         Text            =   "115200"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   3
         Text            =   "500"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Text            =   "1"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "W_Pattern"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   35
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Channel"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   32
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "ComBaud"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Delayms"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "SN_Len"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   29
         Top             =   2160
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   3240
      TabIndex        =   27
      Top             =   600
      Width           =   2295
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "COOL_2"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "COOL_1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   "NORMAL"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "WARM_1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         Caption         =   "WARM_2"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   9
         Top             =   2760
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "ModeIndex"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   5640
      TabIndex        =   21
      Top             =   600
      Width           =   2535
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   13
         Text            =   "2"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   12
         Text            =   "1"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   10
         Text            =   "NA"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   11
         Text            =   "0"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   14
         Text            =   "NA"
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "WARM_1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "NORMAL"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "COOL_2"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "COOL_1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "WARM_2"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   2760
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   8280
      TabIndex        =   20
      Top             =   600
      Width           =   2535
      Begin VB.CheckBox Check10 
         Alignment       =   1  'Right Justify
         Caption         =   "SensorLight"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   18
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CheckBox Check9 
         Alignment       =   1  'Right Justify
         Caption         =   "AdjustOffset"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox Check7 
         Alignment       =   1  'Right Justify
         Caption         =   "CheckColor"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox Check6 
         Alignment       =   1  'Right Justify
         Caption         =   "Save Data"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Value           =   1  'Checked
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   555
      Left            =   8400
      TabIndex        =   19
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   555
      Left            =   9720
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   33
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "frmSetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    sqlstring = "select * from CheckItem where Mark='" & strCurrentModelName & "'"
    Executesql (sqlstring)

    Label1 = strCurrentModelName
 
    Text1.Text = rs("ComBaud")
    Text2.Text = rs("Channel")
    Text3.Text = rs("Delayms")
    Text4.Text = rs("SN_Len")
    Text5.Text = rs("WhitePattern")

    If rs("COOL_2") Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If

    If rs("COOL_1") Then
        Check2.Value = 1
    Else
        Check2.Value = 0
    End If

    If rs("NORMAL") Then
        Check3.Value = 1
    Else
        Check3.Value = 0
    End If

    If rs("WARM_1") Then
        Check4.Value = 1
    Else
        Check4.Value = 0
    End If

    If rs("WARM_2") Then
        Check5.Value = 1
    Else
        Check5.Value = 0
    End If

    If rs("SaveData") Then
        Check6.Value = 1
    Else
        Check6.Value = 0
    End If

    If rs("CheckColor") Then
        Check7.Value = 1
    Else
        Check7.Value = 0
    End If

    If rs("AdjustOFF") Then
        Check9.Value = 1
    Else
        Check9.Value = 0
    End If

    If rs("SensorL") Then
        Check10.Value = 1
    Else
        Check10.Value = 0
    End If

    Text7.Text = rs("Cool_1MI")
    Text6.Text = rs("Cool_2MI")
    Text8.Text = rs("NormalMI")
    Text9.Text = rs("Warm_1MI")
    Text10.Text = rs("Warm_2MI")

    Set rs = Nothing
    Set cn = Nothing
    sqlstring = ""
End Sub

Private Sub Command1_Click()
    sqlstring = "select * from CheckItem where Mark='" & strCurrentModelName & "'"
    Executesql (sqlstring)
  
    rs.Fields(1) = Val(Text1.Text)
    rs.Fields(2) = Val(Text2.Text)
    rs.Fields(3) = Val(Text3.Text)
    rs.Fields(4) = Val(Text5.Text)

    If Check1.Value = 1 Then rs.Fields(5) = True
    If Check1.Value = 0 Then rs.Fields(5) = False
    If Check2.Value = 1 Then rs.Fields(6) = True
    If Check2.Value = 0 Then rs.Fields(6) = False
    If Check3.Value = 1 Then rs.Fields(7) = True
    If Check3.Value = 0 Then rs.Fields(7) = False
    If Check4.Value = 1 Then rs.Fields(8) = True
    If Check4.Value = 0 Then rs.Fields(8) = False
    If Check5.Value = 1 Then rs.Fields(9) = True
    If Check5.Value = 0 Then rs.Fields(9) = False
  
    rs.Fields(10) = Val(Text4.Text)
 
    If Check6.Value = 1 Then rs.Fields(11) = True
    If Check6.Value = 0 Then rs.Fields(11) = False
    If Check7.Value = 1 Then rs.Fields(12) = True
    If Check7.Value = 0 Then rs.Fields(12) = False
    If Check9.Value = 1 Then rs.Fields(13) = True
    If Check9.Value = 0 Then rs.Fields(13) = False
    If Check10.Value = 1 Then rs.Fields(14) = True
    If Check10.Value = 0 Then rs.Fields(14) = False

    rs.Fields(15) = Val(Text6.Text)
    rs.Fields(16) = Val(Text7.Text)
    rs.Fields(17) = Val(Text8.Text)
    rs.Fields(18) = Val(Text9.Text)
    rs.Fields(19) = Val(Text10.Text)

    rs.Update

    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""

    MsgBox "Save success!", vbOKOnly, "warning"
    Unload Me
    Unload Form1
End Sub

