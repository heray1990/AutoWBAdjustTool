VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCmbType 
   Caption         =   "Chroma"
   ClientHeight    =   7935
   ClientLeft      =   6750
   ClientTop       =   4020
   ClientWidth     =   6285
   LinkTopic       =   "Form4"
   ScaleHeight     =   7935
   ScaleWidth      =   6285
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   3840
   End
   Begin VB.ComboBox cmbModel 
      Height          =   300
      ItemData        =   "VPGDemo.frx":0000
      Left            =   360
      List            =   "VPGDemo.frx":002E
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Command"
      Height          =   1335
      Left            =   120
      TabIndex        =   35
      Top             =   600
      Width           =   3135
      Begin VB.CommandButton cmdR 
         Caption         =   "R"
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdG 
         Caption         =   "G"
         Height          =   375
         Left            =   1080
         TabIndex        =   40
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdB 
         Caption         =   "B"
         Height          =   375
         Left            =   2040
         TabIndex        =   39
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdREV 
         Caption         =   "Rev"
         Height          =   375
         Left            =   2040
         TabIndex        =   38
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdOut 
         Caption         =   "Out"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Height          =   375
         Left            =   1080
         TabIndex        =   36
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame grpPenRGB 
      Caption         =   "Change Pen RGB"
      Height          =   2775
      Left            =   120
      TabIndex        =   25
      Top             =   2040
      Width           =   2295
      Begin VB.TextBox txtPenNum 
         Height          =   375
         Left            =   960
         TabIndex        =   30
         Text            =   "0"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtRValue 
         Height          =   375
         Left            =   960
         TabIndex        =   29
         Text            =   "0"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtGValue 
         Height          =   375
         Left            =   960
         TabIndex        =   28
         Text            =   "0"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtBValue 
         Height          =   375
         Left            =   960
         TabIndex        =   27
         Text            =   "0"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdExe 
         Caption         =   "Execute"
         Height          =   375
         Left            =   600
         TabIndex        =   26
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pen No."
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "R"
         Height          =   195
         Left            =   600
         TabIndex        =   33
         Top             =   840
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "G"
         Height          =   195
         Left            =   600
         TabIndex        =   32
         Top             =   1320
         Width           =   120
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "B"
         Height          =   195
         Left            =   600
         TabIndex        =   31
         Top             =   1800
         Width           =   105
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Run"
      Height          =   1335
      Left            =   3480
      TabIndex        =   21
      Top             =   600
      Width           =   2655
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "VPGDemo.frx":009B
         Left            =   120
         List            =   "VPGDemo.frx":00A8
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtRunNum 
         Height          =   375
         Left            =   1560
         TabIndex        =   23
         Text            =   "1"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "Output"
         Height          =   375
         Left            =   1440
         TabIndex        =   22
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Download BMP"
      Height          =   1335
      Left            =   2520
      TabIndex        =   15
      Top             =   2040
      Width           =   3615
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         Height          =   375
         Left            =   2760
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtBmpPath 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   2535
      End
      Begin VB.OptionButton rdoDownload 
         Caption         =   "Download"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   780
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton rdoUpload 
         Caption         =   "Upload"
         Height          =   375
         Left            =   1320
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "GO"
         Height          =   375
         Left            =   2760
         TabIndex        =   16
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "ASCII Command"
      Height          =   855
      Left            =   2520
      TabIndex        =   12
      Top             =   3600
      Width           =   3615
      Begin VB.TextBox txtASCII 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Send"
         Height          =   495
         Left            =   2760
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "A222917"
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   4320
         TabIndex        =   10
         Text            =   "1"
         Top             =   720
         Width           =   735
      End
      Begin VB.DriveListBox Drive1 
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.DirListBox Dir1 
         Height          =   930
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   3015
      End
      Begin VB.FileListBox File1 
         Height          =   990
         Left            =   120
         Pattern         =   "*.PCBA"
         TabIndex        =   7
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox readPath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   4695
      End
      Begin VB.CommandButton btnEnable 
         Caption         =   "Enable"
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton btnRun 
         Caption         =   "Run"
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton btnRep 
         Caption         =   "Report"
         Height          =   375
         Left            =   5040
         TabIndex        =   3
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Block No."
         Height          =   195
         Left            =   3480
         TabIndex        =   11
         Top             =   720
         Width           =   705
      End
   End
   Begin VB.CommandButton btnDwn 
      Caption         =   "Download"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   6960
      Width           =   855
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   3120
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
Dim m_ROpen As Boolean
Dim m_BOpen As Boolean
Dim m_GOpen As Boolean
Dim m_RevOpen As Boolean
Dim m_Status As Boolean
Dim m_IsStop As Boolean
Dim m_StepCnt As Integer
Dim m_Title As String


Public WithEvents Obj As VPGCtrl.VPGCtrl
Attribute Obj.VB_VarHelpID = -1

Private Sub btnDwn_Click()
    Dim parser As VPGParser.IParser
    Dim nRetData() As String
    Dim cmdList() As Long
    Dim blockNo As Integer
    Dim path As String

    Set parser = New VPGParser.VPGParser_A222917
    blockNo = Combo1.Text
    path = readPath.Text
    m_Status = False
    If parser.ReadXML(path, cmdList, nRetData) = True Then
        If ivpg.Download_3(nRetData, blockNo - 1, EnumCmdType_Load) Then
            m_Status = True
        End If
    End If
    
    If m_Status Then
            MsgBox "Download Program Data Complete !"
        Else
            MsgBox "Download Program Data Fail !"
    End If
End Sub

Private Sub btnEnable_Click()
    Dim cmdData(1) As Byte
        cmdData(0) = 0
        cmdData(1) = VPG_CMD_CM_INIT_BOX

        Dim status As Boolean
        status = False
        'If ivpg.Download_2(cmdData, , EnumCmdType_InitilizeCmd) = True Then
            'status = True
        'End If

        If status Then
            MsgBox "Box Enable Successfully !"
        Else
            MsgBox "Box Enable Fail !"
        End If
End Sub

Private Sub btnRep_Click()
    Dim blockNo As Integer
    Dim cmdData(2) As Byte
    Dim cmdIdx As Integer
    Dim seqLenCollection() As Long
    Dim status As Boolean
    Dim RetSeqData() As Byte
    
    status = False
    blockNo = Val(Combo1.Text)
    cmdIdx = 0
    
    cmdData(cmdIdx) = blockNo And &HFF
    cmdIdx = cmdIdx + 1
    cmdData(cmdIdx) = m_StepCnt And &HFF
    cmdIdx = cmdIdx + 1
    cmdData(cmdIdx) = (m_StepCnt And &HFF00) \ 256
    cmdIdx = cmdIdx + 1
    

    If ivpg.Report(cmdData, seqLenCollection, RetSeqData) = True Then
            status = True
    End If

    If status Then
        MsgBox "Report Successfully !"
    Else
        MsgBox "Report Fail !"
    End If
End Sub

Private Sub btnRun_Click()
    Dim cmdData(0) As Byte
        cmdData(0) = Combo1.Text
    Dim RetData() As Byte

    m_Status = False
    If ivpg.StartTest(cmdData, RetData) = True Then
        m_StepCnt = RetData(0) + (RetData(1) * 256)
        m_Status = True
    End If
        
    If m_Status Then
        MsgBox "Run Program Successfully !"
    Else
        MsgBox "Run Program Fail !"
    End If
End Sub

Private Sub btnStop_Click()
    m_IsStop = True
    Call StopTest
End Sub

Private Sub StopTest()
    If m_IsStop = True Then
        m_IsStop = False

        Dim status As Boolean
        status = False
            
        If ivpg.StopTest() = True Then
            status = True
        End If

        If status Then
            MsgBox "Stop Program Successfully !"
        Else
            MsgBox "Stop Program Fail !"
        End If
    End If
End Sub

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

Private Sub Obj_OnAllowToStopTest()
    If m_IsStop = True Then
        m_IsStop = False

        Dim status As Boolean
        status = False
            
        If ivpg.StopTest() = True Then
            status = True
        End If

        If status Then
            MsgBox "Stop Program Successfully !"
        Else
            MsgBox "Stop Program Fail !"
        End If
    End If
End Sub

Private Sub Obj_OnChangedConnectState(ByVal bIsConnected As Boolean)
    If bIsConnected = True Then
            Me.Caption = m_Title + " [Connected " + IIf(ivpg.IsHighSpeed, "USB2.0", "USB1.1") + "]"
        Else
            Me.Caption = m_Title + " [Disconnected]"
    End If
End Sub

Private Sub Obj_OnShowMessage(ByVal msg As String)
    Debug.Print (msg)
End Sub

Private Sub cmdB_Click()
    If m_BOpen = False Then
        ivpg.RunKey VPG_KEY_CKEY_B, 1
    Else
        ivpg.RunKey VPG_KEY_CKEY_B, 0
    End If
    m_BOpen = Not m_BOpen
End Sub

Private Sub cmdExe_Click()

    Dim bBuf(7) As Byte
    bBuf(0) = (CInt(txtPenNum.Text) And &HFF00) \ 256
    bBuf(1) = CInt(txtPenNum.Text) And &HFF
    bBuf(2) = (CInt(txtRValue.Text) And &HFF00) \ 256
    bBuf(3) = CInt(txtRValue.Text) And &HFF
    bBuf(4) = (CInt(txtGValue.Text) And &HFF00) \ 256
    bBuf(5) = CInt(txtGValue.Text) And &HFF
    bBuf(6) = (CInt(txtBValue.Text) And &HFF00) \ 256
    bBuf(7) = CInt(txtBValue.Text) And &HFF

    If ivpg.ExecuteCmd(VPG_CMD_CM_DOWNLOAD, VPG_SCMD_SCM_CTL_RUNRGB, bBuf, False) = False Then
        MsgBox "RunRGB Function Perform Fail. "
    End If
    
End Sub

Private Sub cmdG_Click()
    If m_GOpen = False Then
        ivpg.RunKey VPG_KEY_CKEY_G, 1
    Else
        ivpg.RunKey VPG_KEY_CKEY_G, 0
    End If
    m_GOpen = Not m_GOpen
End Sub

Private Sub cmdOpen_Click()
    Dim strfilter As String
    
    strfilter = "Bitmap File (*.bmp)|*.bmp"
    If rdoDownload.value = True Then
        dlgOpen.ShowOpen
    Else
        dlgOpen.ShowSave
    End If
    
    dlgOpen.Filter = strfilter
    dlgOpen.InitDir = "..\"
    
    If dlgOpen.FileName <> "" Then
            txtBmpPath.Text = dlgOpen.FileName
    End If
End Sub

Private Sub cmdOut_Click()
    ivpg.RunKey (VPG_KEY_CKEY_OUT)
    m_ROpen = True
    m_GOpen = True
    m_BOpen = True
    m_RevOpen = False
End Sub

Private Sub cmdQuit_Click()
    ivpg.RunKey (VPG_KEY_CKEY_QUIT)
    m_ROpen = False
    m_GOpen = False
    m_BOpen = False
    m_RevOpen = False
End Sub

Private Sub cmdR_Click()
    If m_ROpen = False Then
        ivpg.RunKey VPG_KEY_CKEY_R, 1
    Else
        ivpg.RunKey VPG_KEY_CKEY_R, 0
    End If
    m_ROpen = Not m_ROpen
End Sub

Private Sub cmdREV_Click()
    If m_RevOpen = False Then
        ivpg.RunKey VPG_KEY_CKEY_REV, 1
    Else
        ivpg.RunKey VPG_KEY_CKEY_REV, 0
    End If
    m_RevOpen = Not m_RevOpen
   
End Sub

Private Sub cmdRun_Click()

    Dim bNo(1) As Byte
    bNo(0) = (CInt(txtRunNum.Text) And &HFF00) \ 256
    bNo(1) = CInt(txtRunNum.Text) And &HFF
   
    Select Case cmbType.Text
        Case "Timing"
            ivpg.ExecuteCmd VPG_CMD_CM_DOWNLOAD, VPG_SCMD_SCM_CTL_RUNTIM, bNo, False
        Case "Pattern"
            ivpg.RunKey (VPG_KEY_CKEY_OUT)
            ivpg.ExecuteCmd VPG_CMD_CM_DOWNLOAD, VPG_SCMD_SCM_CTL_RUNPTN, bNo, False
        Case "Program"
            Select Case cmbModel.Text
                Case "2401"
                Case "2402"
                Case "2333_B"
                Case "23293_B"
                Case "23294"
                    ivpg.RunKey (VPG_KEY_CKEY_QUIT)
                    If ivpg.RunKey(VPG_KEY_CKEY_PRG) Then
                            Dim bData(1) As Byte

                            If ivpg.ExecuteCmd(VPG_CMD_CM_DOWNLOAD, VPG_SCMD_SCM_CTL_PRGL, bNo, False) Then
                                If ivpg.RunKey(VPG_KEY_CKEY_OUT) Then

                                End If
                            End If
                        End If

                Case Else
                    ivpg.ExecuteCmd VPG_CMD_CM_DOWNLOAD, VPG_SCMD_SCM_CTL_RUNTIMPTN, bNo, False
            End Select
            
    End Select
End Sub





Private Sub Dir1_Change()
    File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
    Dir1.path = Drive1.Drive
End Sub

Private Sub File1_Click()
    Dim pathlen As Integer
    pathlen = Len(Dir1.path)
    If (pathlen = 3) Then
        readPath.Text = Dir1.path + File1.FileName
    Else
        readPath.Text = Dir1.path + "\" + File1.FileName
    End If
End Sub

Private Sub Form_Load()
    
    Dim i As Integer
    
    m_Title = Me.Caption
    cmbModel.Text = "22294_A"
    
    cmbType.Text = "Pattern"

    mTitle = Me.Caption

    For i = 1 To 50
        Combo1.AddItem (i)
    Next

    txtRunNum.Text = "103"
    cmbModel_Click

End Sub





Private Sub Command1_Click()
    Dim file_num
    Dim byteUploadData() As Byte
    Dim byteDownloadData() As Byte

    file_num = FreeFile
    
    If txtBmpPath.Text = "" Then
        Exit Sub
    End If
    

    
    
    
    If rdoDownload.value = True Then
        Open txtBmpPath.Text For Binary As #file_num
        ReDim byteDownloadData(FileLen(txtBmpPath.Text))
        Dim i As Long
        i = 0
        Do While Not EOF(file_num)
            Get #file_num, , byteDownloadData(i)
            i = i + 1
        Loop
        Close #file_num
        
        If ivpg.Download(VPG_DATA_BMP, byteDownloadData, 1, 1) = False Then
            MsgBox ("Download Bitmap Data Fail !")
        Else
            MsgBox ("Download Bitmap Complete !")
        End If
        
    Else
        If ivpg.Upload(VPG_DATA_BMP, byteUploadData, 1) = False Then
            MsgBox ("Upload Bitmap Data Fail !")
        Else
            Open txtBmpPath.Text For Binary As #file_num
            Put #file_num, , byteUploadData
            Close #file_num
            MsgBox ("Upload Bitmap Complete !")
        End If
    
        ivpg.RunKey (VPG_KEY_CKEY_QUIT)
    End If
    
    
      
  

   
    
    
    
    
End Sub

Private Sub Command2_Click()

    Dim i As Integer
    If txtASCII.Text = "" Or Len(txtASCII.Text) > &H100 Then
        Exit Sub
    End If
    
  
    Dim byteASCIIData() As Byte
      
    ReDim byteASCIIData(Len(txtASCII.Text))
      
    For i = 0 To Len(txtASCII.Text) - 1
        byteASCIIData(i + 1) = Asc(Mid(txtASCII.Text, i + 1, 1))
    Next
    

    byteASCIIData(0) = Len(txtASCII.Text)
      
    
    ivpg.ExecuteCmd_2 VPG_CMD_CM_DOWNLOAD, VPG_SCMD_SCM_ANSII_CMD, byteASCIIData, False
    
   
    
End Sub

Private Sub OnChangedConnectState(ByVal bIsConnected As Boolean)
    
    
    If bIsConnected = True Then
        Form1.Caption = mTitle + " [Connected " + IIf(ivpg.IsHighSpeed, "USB2.0", "USB1.1") + "]"
    Else
        Form1.Caption = mTitle + " [Disconnected]"
    End If
    
End Sub




Private Sub rdoDownload_Click()
    cmdOpen.Caption = "Open"
End Sub

Private Sub rdoUpload_Click()
    cmdOpen.Caption = "Save"
End Sub

Private Sub txtBValue_Validate(Cancel As Boolean)
    If txtBValue.Text = "" Then
        Cancel = True
    ElseIf IsNumeric(txtBValue.Text) = False Then
        Cancel = True
    ElseIf Val(txtBValue.Text) < 0 Or Val(txtBValue.Text) > 4095 Then
        Cancel = True
    End If
        
        
    
End Sub

Private Sub txtGValue_Validate(Cancel As Boolean)
    If txtGValue.Text = "" Then
        Cancel = True
    ElseIf IsNumeric(txtGValue.Text) = False Then
        Cancel = True
    ElseIf Val(txtGValue.Text) < 0 Or Val(txtGValue.Text) > 4095 Then
        Cancel = True
    End If
End Sub



Private Sub txtPenNum_Validate(Cancel As Boolean)
        If txtPenNum.Text = "" Then
        Cancel = True
    ElseIf IsNumeric(txtPenNum.Text) = False Then
        Cancel = True
    ElseIf Val(txtPenNum.Text) < 0 Or Val(txtPenNum.Text) > 1023 Then
        Cancel = True
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


Public Sub ChangePattern(Ptn As String)
    Dim bNo(1) As Byte

    cmbType.Text = "Pattern"
    txtRunNum.Text = Ptn
    
    bNo(0) = (CInt(txtRunNum.Text) And &HFF00) \ 256
    bNo(1) = CInt(txtRunNum.Text) And &HFF

    ivpg.RunKey (VPG_KEY_CKEY_OUT)
    ivpg.ExecuteCmd VPG_CMD_CM_DOWNLOAD, VPG_SCMD_SCM_CTL_RUNPTN, bNo, False
End Sub
