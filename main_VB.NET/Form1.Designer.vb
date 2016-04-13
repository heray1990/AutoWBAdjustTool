<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form1
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents vbConCA310 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents tbDisConnectastro As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents vbFunc As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents vbSetSPEC As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents vbSet As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents vbAbout As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents vbDescription As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents tcpClient As AxMSWinsockLib.AxWinsock
	Public WithEvents lbColorTempWrong As System.Windows.Forms.Label
	Public WithEvents Picture1 As System.Windows.Forms.Panel
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents MSComm1 As AxMSCommLib.AxMSComm
	Public WithEvents CheckStep As System.Windows.Forms.TextBox
	Public WithEvents txtInput As System.Windows.Forms.TextBox
	Public WithEvents lbCommMode As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents lbModelName As System.Windows.Forms.Label
	Public WithEvents lbTimer As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents lbAdjustWARM_2 As System.Windows.Forms.Label
	Public WithEvents lbAdjustCOOL_2 As System.Windows.Forms.Label
	Public WithEvents Label_Lv As System.Windows.Forms.Label
	Public WithEvents Label_y As System.Windows.Forms.Label
	Public WithEvents Label_x As System.Windows.Forms.Label
	Public WithEvents lbAdjustWARM_1 As System.Windows.Forms.Label
	Public WithEvents lbAdjustNormal As System.Windows.Forms.Label
	Public WithEvents lbAdjustCOOL_1 As System.Windows.Forms.Label
	Public WithEvents checkResult As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.vbFunc = New System.Windows.Forms.ToolStripMenuItem
        Me.vbConCA310 = New System.Windows.Forms.ToolStripMenuItem
        Me.tbDisConnectastro = New System.Windows.Forms.ToolStripMenuItem
        Me.vbSet = New System.Windows.Forms.ToolStripMenuItem
        Me.vbSetSPEC = New System.Windows.Forms.ToolStripMenuItem
        Me.vbDescription = New System.Windows.Forms.ToolStripMenuItem
        Me.vbAbout = New System.Windows.Forms.ToolStripMenuItem
        Me.tcpClient = New AxMSWinsockLib.AxWinsock
        Me.Picture1 = New System.Windows.Forms.Panel
        Me.lbColorTempWrong = New System.Windows.Forms.Label
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.MSComm1 = New AxMSCommLib.AxMSComm
        Me.CheckStep = New System.Windows.Forms.TextBox
        Me.txtInput = New System.Windows.Forms.TextBox
        Me.lbCommMode = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.lbModelName = New System.Windows.Forms.Label
        Me.lbTimer = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.lbAdjustWARM_2 = New System.Windows.Forms.Label
        Me.lbAdjustCOOL_2 = New System.Windows.Forms.Label
        Me.Label_Lv = New System.Windows.Forms.Label
        Me.Label_y = New System.Windows.Forms.Label
        Me.Label_x = New System.Windows.Forms.Label
        Me.lbAdjustWARM_1 = New System.Windows.Forms.Label
        Me.lbAdjustNormal = New System.Windows.Forms.Label
        Me.lbAdjustCOOL_1 = New System.Windows.Forms.Label
        Me.checkResult = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.MainMenu1.SuspendLayout()
        CType(Me.tcpClient, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Picture1.SuspendLayout()
        CType(Me.MSComm1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.vbFunc, Me.vbSet, Me.vbDescription})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(689, 24)
        Me.MainMenu1.TabIndex = 22
        '
        'vbFunc
        '
        Me.vbFunc.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.vbConCA310, Me.tbDisConnectastro})
        Me.vbFunc.Name = "vbFunc"
        Me.vbFunc.Size = New System.Drawing.Size(65, 20)
        Me.vbFunc.Text = "Function"
        '
        'vbConCA310
        '
        Me.vbConCA310.Name = "vbConCA310"
        Me.vbConCA310.Size = New System.Drawing.Size(178, 22)
        Me.vbConCA310.Text = "ConnectCA210"
        '
        'tbDisConnectastro
        '
        Me.tbDisConnectastro.Name = "tbDisConnectastro"
        Me.tbDisConnectastro.Size = New System.Drawing.Size(178, 22)
        Me.tbDisConnectastro.Text = "DisConnectCA210(&D)"
        '
        'vbSet
        '
        Me.vbSet.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.vbSetSPEC})
        Me.vbSet.Name = "vbSet"
        Me.vbSet.Size = New System.Drawing.Size(59, 20)
        Me.vbSet.Text = "Setting"
        '
        'vbSetSPEC
        '
        Me.vbSetSPEC.Name = "vbSetSPEC"
        Me.vbSetSPEC.Size = New System.Drawing.Size(118, 22)
        Me.vbSetSPEC.Text = "Set Spec"
        '
        'vbDescription
        '
        Me.vbDescription.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.vbAbout})
        Me.vbDescription.Name = "vbDescription"
        Me.vbDescription.Size = New System.Drawing.Size(83, 20)
        Me.vbDescription.Text = "Description"
        '
        'vbAbout
        '
        Me.vbAbout.Name = "vbAbout"
        Me.vbAbout.ShortcutKeys = System.Windows.Forms.Keys.F2
        Me.vbAbout.Size = New System.Drawing.Size(117, 22)
        Me.vbAbout.Text = "About"
        '
        'tcpClient
        '
        Me.tcpClient.Enabled = True
        Me.tcpClient.Location = New System.Drawing.Point(704, 224)
        Me.tcpClient.Name = "tcpClient"
        Me.tcpClient.OcxState = CType(resources.GetObject("tcpClient.OcxState"), System.Windows.Forms.AxHost.State)
        Me.tcpClient.Size = New System.Drawing.Size(28, 28)
        Me.tcpClient.TabIndex = 0
        '
        'Picture1
        '
        Me.Picture1.BackColor = System.Drawing.SystemColors.Control
        Me.Picture1.BackgroundImage = CType(resources.GetObject("Picture1.BackgroundImage"), System.Drawing.Image)
        Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Picture1.Controls.Add(Me.lbColorTempWrong)
        Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Picture1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Picture1.Location = New System.Drawing.Point(176, 88)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(254, 172)
        Me.Picture1.TabIndex = 6
        Me.Picture1.TabStop = True
        '
        'lbColorTempWrong
        '
        Me.lbColorTempWrong.BackColor = System.Drawing.Color.Transparent
        Me.lbColorTempWrong.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbColorTempWrong.ForeColor = System.Drawing.Color.Red
        Me.lbColorTempWrong.Location = New System.Drawing.Point(24, 1)
        Me.lbColorTempWrong.Name = "lbColorTempWrong"
        Me.lbColorTempWrong.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbColorTempWrong.Size = New System.Drawing.Size(65, 17)
        Me.lbColorTempWrong.TabIndex = 7
        Me.lbColorTempWrong.Text = "Out Range"
        Me.lbColorTempWrong.Visible = False
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'MSComm1
        '
        Me.MSComm1.Enabled = True
        Me.MSComm1.Location = New System.Drawing.Point(704, 288)
        Me.MSComm1.Name = "MSComm1"
        Me.MSComm1.OcxState = CType(resources.GetObject("MSComm1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.MSComm1.Size = New System.Drawing.Size(38, 38)
        Me.MSComm1.TabIndex = 7
        '
        'CheckStep
        '
        Me.CheckStep.AcceptsReturn = True
        Me.CheckStep.BackColor = System.Drawing.SystemColors.Control
        Me.CheckStep.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CheckStep.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.CheckStep.Font = New System.Drawing.Font("Arial", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckStep.ForeColor = System.Drawing.SystemColors.WindowText
        Me.CheckStep.Location = New System.Drawing.Point(429, 88)
        Me.CheckStep.MaxLength = 0
        Me.CheckStep.Multiline = True
        Me.CheckStep.Name = "CheckStep"
        Me.CheckStep.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CheckStep.Size = New System.Drawing.Size(254, 239)
        Me.CheckStep.TabIndex = 5
        Me.CheckStep.Text = "CheckStep" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'txtInput
        '
        Me.txtInput.AcceptsReturn = True
        Me.txtInput.BackColor = System.Drawing.SystemColors.Window
        Me.txtInput.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtInput.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtInput.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInput.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInput.Location = New System.Drawing.Point(8, 64)
        Me.txtInput.MaxLength = 0
        Me.txtInput.Name = "txtInput"
        Me.txtInput.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtInput.Size = New System.Drawing.Size(169, 25)
        Me.txtInput.TabIndex = 1
        Me.txtInput.Text = "123456789"
        Me.txtInput.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lbCommMode
        '
        Me.lbCommMode.BackColor = System.Drawing.SystemColors.Control
        Me.lbCommMode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbCommMode.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbCommMode.Font = New System.Drawing.Font("Arial", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbCommMode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lbCommMode.Location = New System.Drawing.Point(8, 88)
        Me.lbCommMode.Name = "lbCommMode"
        Me.lbCommMode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbCommMode.Size = New System.Drawing.Size(169, 35)
        Me.lbCommMode.TabIndex = 21
        Me.lbCommMode.Text = "UART"
        Me.lbCommMode.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(300, 295)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(65, 30)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "2970"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lbModelName
        '
        Me.lbModelName.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lbModelName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbModelName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbModelName.Font = New System.Drawing.Font("Arial", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbModelName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lbModelName.Location = New System.Drawing.Point(8, 24)
        Me.lbModelName.Name = "lbModelName"
        Me.lbModelName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbModelName.Size = New System.Drawing.Size(169, 41)
        Me.lbModelName.TabIndex = 19
        Me.lbModelName.Text = "Sampl1"
        Me.lbModelName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lbTimer
        '
        Me.lbTimer.BackColor = System.Drawing.SystemColors.Control
        Me.lbTimer.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbTimer.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbTimer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lbTimer.Location = New System.Drawing.Point(376, 295)
        Me.lbTimer.Name = "lbTimer"
        Me.lbTimer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbTimer.Size = New System.Drawing.Size(50, 30)
        Me.lbTimer.TabIndex = 18
        Me.lbTimer.Text = "0s"
        Me.lbTimer.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(205, 295)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(65, 30)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "2670"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label7.Location = New System.Drawing.Point(176, 292)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(254, 35)
        Me.Label7.TabIndex = 16
        Me.Label7.Text = "SPEC"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label6.Location = New System.Drawing.Point(8, 292)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(169, 35)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "INITIAL"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lbAdjustWARM_2
        '
        Me.lbAdjustWARM_2.BackColor = System.Drawing.SystemColors.Control
        Me.lbAdjustWARM_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbAdjustWARM_2.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbAdjustWARM_2.Font = New System.Drawing.Font("Arial", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbAdjustWARM_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lbAdjustWARM_2.Location = New System.Drawing.Point(8, 258)
        Me.lbAdjustWARM_2.Name = "lbAdjustWARM_2"
        Me.lbAdjustWARM_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbAdjustWARM_2.Size = New System.Drawing.Size(169, 35)
        Me.lbAdjustWARM_2.TabIndex = 14
        Me.lbAdjustWARM_2.Text = "WARM2"
        Me.lbAdjustWARM_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lbAdjustCOOL_2
        '
        Me.lbAdjustCOOL_2.BackColor = System.Drawing.SystemColors.Control
        Me.lbAdjustCOOL_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbAdjustCOOL_2.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbAdjustCOOL_2.Font = New System.Drawing.Font("Arial", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbAdjustCOOL_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lbAdjustCOOL_2.Location = New System.Drawing.Point(8, 156)
        Me.lbAdjustCOOL_2.Name = "lbAdjustCOOL_2"
        Me.lbAdjustCOOL_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbAdjustCOOL_2.Size = New System.Drawing.Size(169, 35)
        Me.lbAdjustCOOL_2.TabIndex = 13
        Me.lbAdjustCOOL_2.Text = "COOL2"
        Me.lbAdjustCOOL_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_Lv
        '
        Me.Label_Lv.BackColor = System.Drawing.SystemColors.Window
        Me.Label_Lv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label_Lv.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label_Lv.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Lv.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label_Lv.Location = New System.Drawing.Point(366, 259)
        Me.Label_Lv.Name = "Label_Lv"
        Me.Label_Lv.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label_Lv.Size = New System.Drawing.Size(64, 34)
        Me.Label_Lv.TabIndex = 12
        Me.Label_Lv.Text = "210"
        Me.Label_Lv.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_y
        '
        Me.Label_y.BackColor = System.Drawing.SystemColors.Window
        Me.Label_y.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label_y.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_y.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label_y.Location = New System.Drawing.Point(300, 261)
        Me.Label_y.Name = "Label_y"
        Me.Label_y.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label_y.Size = New System.Drawing.Size(65, 30)
        Me.Label_y.TabIndex = 11
        Me.Label_y.Text = "2800"
        Me.Label_y.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_x
        '
        Me.Label_x.BackColor = System.Drawing.SystemColors.Window
        Me.Label_x.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label_x.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_x.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label_x.Location = New System.Drawing.Point(205, 261)
        Me.Label_x.Name = "Label_x"
        Me.Label_x.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label_x.Size = New System.Drawing.Size(65, 30)
        Me.Label_x.TabIndex = 10
        Me.Label_x.Text = "2700"
        Me.Label_x.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lbAdjustWARM_1
        '
        Me.lbAdjustWARM_1.BackColor = System.Drawing.SystemColors.Control
        Me.lbAdjustWARM_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbAdjustWARM_1.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbAdjustWARM_1.Font = New System.Drawing.Font("Arial", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbAdjustWARM_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lbAdjustWARM_1.Location = New System.Drawing.Point(8, 224)
        Me.lbAdjustWARM_1.Name = "lbAdjustWARM_1"
        Me.lbAdjustWARM_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbAdjustWARM_1.Size = New System.Drawing.Size(169, 35)
        Me.lbAdjustWARM_1.TabIndex = 4
        Me.lbAdjustWARM_1.Text = "WARM1"
        Me.lbAdjustWARM_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lbAdjustNormal
        '
        Me.lbAdjustNormal.BackColor = System.Drawing.SystemColors.Control
        Me.lbAdjustNormal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbAdjustNormal.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbAdjustNormal.Font = New System.Drawing.Font("Arial", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbAdjustNormal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lbAdjustNormal.Location = New System.Drawing.Point(8, 190)
        Me.lbAdjustNormal.Name = "lbAdjustNormal"
        Me.lbAdjustNormal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbAdjustNormal.Size = New System.Drawing.Size(169, 35)
        Me.lbAdjustNormal.TabIndex = 3
        Me.lbAdjustNormal.Text = "NORMAL"
        Me.lbAdjustNormal.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lbAdjustCOOL_1
        '
        Me.lbAdjustCOOL_1.BackColor = System.Drawing.SystemColors.Control
        Me.lbAdjustCOOL_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbAdjustCOOL_1.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbAdjustCOOL_1.Font = New System.Drawing.Font("Arial", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbAdjustCOOL_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lbAdjustCOOL_1.Location = New System.Drawing.Point(8, 122)
        Me.lbAdjustCOOL_1.Name = "lbAdjustCOOL_1"
        Me.lbAdjustCOOL_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbAdjustCOOL_1.Size = New System.Drawing.Size(169, 35)
        Me.lbAdjustCOOL_1.TabIndex = 2
        Me.lbAdjustCOOL_1.Text = "COOL1"
        Me.lbAdjustCOOL_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'checkResult
        '
        Me.checkResult.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.checkResult.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.checkResult.Cursor = System.Windows.Forms.Cursors.Default
        Me.checkResult.Font = New System.Drawing.Font("Arial", 38.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.checkResult.ForeColor = System.Drawing.SystemColors.WindowText
        Me.checkResult.Location = New System.Drawing.Point(176, 24)
        Me.checkResult.Name = "checkResult"
        Me.checkResult.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.checkResult.Size = New System.Drawing.Size(507, 65)
        Me.checkResult.TabIndex = 0
        Me.checkResult.Text = " ADJUST COLOR"
        Me.checkResult.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Window
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.Location = New System.Drawing.Point(176, 259)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(96, 34)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "x:"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Window
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label4.Location = New System.Drawing.Point(271, 259)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(96, 34)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "y:"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(689, 332)
        Me.Controls.Add(Me.tcpClient)
        Me.Controls.Add(Me.Picture1)
        Me.Controls.Add(Me.MSComm1)
        Me.Controls.Add(Me.CheckStep)
        Me.Controls.Add(Me.txtInput)
        Me.Controls.Add(Me.lbCommMode)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lbModelName)
        Me.Controls.Add(Me.lbTimer)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.lbAdjustWARM_2)
        Me.Controls.Add(Me.lbAdjustCOOL_2)
        Me.Controls.Add(Me.Label_Lv)
        Me.Controls.Add(Me.Label_y)
        Me.Controls.Add(Me.Label_x)
        Me.Controls.Add(Me.lbAdjustWARM_1)
        Me.Controls.Add(Me.lbAdjustNormal)
        Me.Controls.Add(Me.lbAdjustCOOL_1)
        Me.Controls.Add(Me.checkResult)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(391, 175)
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Auto Color Temp Adjust System"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        CType(Me.tcpClient, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Picture1.ResumeLayout(False)
        CType(Me.MSComm1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class