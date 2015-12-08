VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12765
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   ScaleHeight     =   5850
   ScaleWidth      =   12765
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   7800
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   9840
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Text            =   "Result"
      Top             =   1080
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   1
      Text            =   "55 17 00 00 00 00 00 00 00 00 E9 FE"
      Top             =   450
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  strBuff = ""
  Text2 = ""
  Call SET_AUTO_ADC
  DelayMS 5000
  Text2 = strBuff
End Sub

Private Sub Command2_Click()
  strBuff = ""
  Text2 = ""
  Call SET_COMMAND_RS
  DelayMS 2000
  Text2 = strBuff
End Sub

Private Sub Command3_Click()
Dim SendDataBuf(0 To 11) As Byte
SendDataBuf(0) = Mid(Text1.Text, 1, 2)
SendDataBuf(1) = Mid(Text1.Text, 4, 2)
SendDataBuf(2) = Mid(Text1.Text, 7, 2)
SendDataBuf(3) = Mid(Text1.Text, 10, 2)
SendDataBuf(4) = Mid(Text1.Text, 13, 2)
SendDataBuf(5) = Mid(Text1.Text, 16, 2)
SendDataBuf(6) = Mid(Text1.Text, 19, 2)
SendDataBuf(7) = Mid(Text1.Text, 22, 2)
SendDataBuf(8) = Mid(Text1.Text, 25, 2)
SendDataBuf(9) = Mid(Text1.Text, 28, 2)
SendDataBuf(10) = Mid(Text1.Text, 31, 2)

 Text2.Text = SendDataBuf(0) & ":" & SendDataBuf(1) & ":" & SendDataBuf(2) & ":" & SendDataBuf(3) & ":" & SendDataBuf(4) & ":" & SendDataBuf(5) & ":" & SendDataBuf(6) & ":" & SendDataBuf(7) & ":" & SendDataBuf(8) & ":" & SendDataBuf(9) & ":" & SendDataBuf(10)
End Sub

