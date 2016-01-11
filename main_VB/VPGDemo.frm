VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCmbType 
   Caption         =   "Chroma"
   ClientHeight    =   1590
   ClientLeft      =   6750
   ClientTop       =   4020
   ClientWidth     =   3420
   Icon            =   "VPGDemo.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   1590
   ScaleWidth      =   3420
   Begin VB.ComboBox cmbModel 
      Height          =   300
      ItemData        =   "VPGDemo.frx":1DF72
      Left            =   360
      List            =   "VPGDemo.frx":1DFA0
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Run"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2655
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "VPGDemo.frx":1E00D
         Left            =   120
         List            =   "VPGDemo.frx":1E01A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtRunNum 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Text            =   "1"
         Top             =   360
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   2400
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   """.bmp"""
      DialogTitle     =   "Selet BMP File"
      Filter          =   "*.bmp"
   End
End
Attribute VB_Name = "frmCmbType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim ivpg As IVPGCtrl
Private mTitle As String
Dim m_Title As String


Public WithEvents Obj As VPGCtrl.VPGCtrl
Attribute Obj.VB_VarHelpID = -1

Private Sub cmbModel_Click()

    Select Case cmbModel.Text
        Case "2401"
            Set ivpg = New VPGCtrl.VPGCtrl_24xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG2401)
        Case "2402"
            Set ivpg = New VPGCtrl.VPGCtrl_24xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG2402)
        Case "2333_B"
            Set ivpg = New VPGCtrl.VPGCtrl_24xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG2333_B)
        Case "23293_B"
            Set ivpg = New VPGCtrl.VPGCtrl_24xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG23293_B)
        Case "23294"
            Set ivpg = New VPGCtrl.VPGCtrl_24xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG23294)
        Case "22293"
            Set ivpg = New VPGCtrl.VPGCtrl_22xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG22293)
        Case "22293_A"
            Set ivpg = New VPGCtrl.VPGCtrl_22xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG22293_A)
        Case "22293_B"
            Set ivpg = New VPGCtrl.VPGCtrl_22xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG22293_B)
        Case "2233"
            Set ivpg = New VPGCtrl.VPGCtrl_22xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG2233)
        Case "2233_A"
            Set ivpg = New VPGCtrl.VPGCtrl_22xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG2233_A)
        Case "2233_B"
            Set ivpg = New VPGCtrl.VPGCtrl_22xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG2233_B)
        Case "2234"
            Set ivpg = New VPGCtrl.VPGCtrl_22xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG2234)
        Case "22294"
            Set ivpg = New VPGCtrl.VPGCtrl_22xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG22294)
        Case "22294_A"
            Set ivpg = New VPGCtrl.VPGCtrl_22xx
            Set Obj = ivpg
            ivpg.InitDevice (VPG_MODEL_VPG22294_A)
    End Select

End Sub

Private Sub Obj_OnChangedConnectState(ByVal bIsConnected As Boolean)
    If bIsConnected = True Then
            Me.Caption = m_Title + " [Connected " + IIf(ivpg.IsHighSpeed, "USB2.0", "USB1.1") + "]"
        Else
            Me.Caption = m_Title + " [Disconnected]"
    End If
End Sub

Private Sub Form_Load()
    
    Dim i As Integer
    
    m_Title = Me.Caption
    cmbModel.Text = "22294_A"
    
    cmbType.Text = "Pattern"

    mTitle = Me.Caption

    txtRunNum.Text = "103"
    cmbModel_Click

End Sub

Private Sub OnChangedConnectState(ByVal bIsConnected As Boolean)
    If bIsConnected = True Then
        Form1.Caption = mTitle + " [Connected " + IIf(ivpg.IsHighSpeed, "USB2.0", "USB1.1") + "]"
    Else
        Form1.Caption = mTitle + " [Disconnected]"
    End If
End Sub


Private Sub txtRunNum_Validate(Cancel As Boolean)
   If txtRunNum.Text = "" Then
        Cancel = True
    ElseIf IsNumeric(txtRunNum.Text) = False Then
        Cancel = True
    Else
        Select Case cmbType.Text
            Case "Timing", "Pattern"
                If Val(txtRunNum.Text) < 0 Or Val(txtRunNum.Text) > 5000 Then
                    Cancel = True
                End If
            Case "Program"
                If Val(txtRunNum.Text) < 0 Or Val(txtRunNum.Text) > 1000 Then
                    Cancel = True
                End If
        End Select
    End If
End Sub

Private Sub txtRValue_Validate(Cancel As Boolean)
    If txtRValue.Text = "" Then
        Cancel = True
    ElseIf IsNumeric(txtRValue.Text) = False Then
        Cancel = True
    ElseIf Val(txtRValue.Text) < 0 Or Val(txtRValue.Text) > 4095 Then
        Cancel = True
    End If
End Sub

Public Sub ChangeTiming(Tim As String)
    Dim bNo(1) As Byte

    cmbType.Text = "Timing"
    txtRunNum.Text = Tim
    
    bNo(0) = (CInt(txtRunNum.Text) And &HFF00) \ 256
    bNo(1) = CInt(txtRunNum.Text) And &HFF

    ivpg.ExecuteCmd VPG_CMD_CM_DOWNLOAD, VPG_SCMD_SCM_CTL_RUNTIM, bNo, False
End Sub

Public Sub ChangePattern(Ptn As String)
    Dim bNo(1) As Byte

    cmbType.Text = "Pattern"
    txtRunNum.Text = Ptn
    
    bNo(0) = (CInt(txtRunNum.Text) And &HFF00) \ 256
    bNo(1) = CInt(txtRunNum.Text) And &HFF

    ivpg.RunKey (VPG_KEY_CKEY_OUT)
    ivpg.ExecuteCmd VPG_CMD_CM_DOWNLOAD, VPG_SCMD_SCM_CTL_RUNPTN, bNo, False
End Sub
