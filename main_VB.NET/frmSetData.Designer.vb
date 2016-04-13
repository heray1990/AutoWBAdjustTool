<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSetData
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
	Public WithEvents cmbChromaModel As System.Windows.Forms.ComboBox
	Public WithEvents lbChromaModel As System.Windows.Forms.Label
	Public WithEvents Frame6 As System.Windows.Forms.GroupBox
	Public WithEvents cmbComID As System.Windows.Forms.ComboBox
	Public WithEvents cmbComBaud As System.Windows.Forms.ComboBox
	Public WithEvents lbComId As System.Windows.Forms.Label
	Public WithEvents lbComBaud As System.Windows.Forms.Label
	Public WithEvents Frame5 As System.Windows.Forms.GroupBox
	Public WithEvents optUart As System.Windows.Forms.RadioButton
	Public WithEvents optNetwork As System.Windows.Forms.RadioButton
	Public WithEvents Frame4 As System.Windows.Forms.GroupBox
	Public WithEvents txtLvSpec As System.Windows.Forms.TextBox
	Public WithEvents cmbInputSource As System.Windows.Forms.ComboBox
	Public WithEvents txtSNLen As System.Windows.Forms.TextBox
	Public WithEvents txtDelay As System.Windows.Forms.TextBox
	Public WithEvents lbLvSpec As System.Windows.Forms.Label
	Public WithEvents lbInputSrc As System.Windows.Forms.Label
	Public WithEvents lbSnLen As System.Windows.Forms.Label
	Public WithEvents lbDelayMs As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents txtChannel As System.Windows.Forms.TextBox
	Public WithEvents lbChannelNum As System.Windows.Forms.Label
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
	Public WithEvents Check7 As System.Windows.Forms.CheckBox
	Public WithEvents Check6 As System.Windows.Forms.CheckBox
	Public WithEvents Check1 As System.Windows.Forms.CheckBox
	Public WithEvents Check2 As System.Windows.Forms.CheckBox
	Public WithEvents Check3 As System.Windows.Forms.CheckBox
	Public WithEvents Check4 As System.Windows.Forms.CheckBox
	Public WithEvents Check5 As System.Windows.Forms.CheckBox
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSetData))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Frame6 = New System.Windows.Forms.GroupBox
		Me.cmbChromaModel = New System.Windows.Forms.ComboBox
		Me.lbChromaModel = New System.Windows.Forms.Label
		Me.Frame5 = New System.Windows.Forms.GroupBox
		Me.cmbComID = New System.Windows.Forms.ComboBox
		Me.cmbComBaud = New System.Windows.Forms.ComboBox
		Me.lbComId = New System.Windows.Forms.Label
		Me.lbComBaud = New System.Windows.Forms.Label
		Me.Frame4 = New System.Windows.Forms.GroupBox
		Me.optUart = New System.Windows.Forms.RadioButton
		Me.optNetwork = New System.Windows.Forms.RadioButton
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me.txtLvSpec = New System.Windows.Forms.TextBox
		Me.cmbInputSource = New System.Windows.Forms.ComboBox
		Me.txtSNLen = New System.Windows.Forms.TextBox
		Me.txtDelay = New System.Windows.Forms.TextBox
		Me.lbLvSpec = New System.Windows.Forms.Label
		Me.lbInputSrc = New System.Windows.Forms.Label
		Me.lbSnLen = New System.Windows.Forms.Label
		Me.lbDelayMs = New System.Windows.Forms.Label
		Me.Frame3 = New System.Windows.Forms.GroupBox
		Me.txtChannel = New System.Windows.Forms.TextBox
		Me.lbChannelNum = New System.Windows.Forms.Label
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.Check7 = New System.Windows.Forms.CheckBox
		Me.Check6 = New System.Windows.Forms.CheckBox
		Me.Check1 = New System.Windows.Forms.CheckBox
		Me.Check2 = New System.Windows.Forms.CheckBox
		Me.Check3 = New System.Windows.Forms.CheckBox
		Me.Check4 = New System.Windows.Forms.CheckBox
		Me.Check5 = New System.Windows.Forms.CheckBox
		Me.Command1 = New System.Windows.Forms.Button
		Me.Label1 = New System.Windows.Forms.Label
		Me.Frame6.SuspendLayout()
		Me.Frame5.SuspendLayout()
		Me.Frame4.SuspendLayout()
		Me.Frame2.SuspendLayout()
		Me.Frame3.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "SpecData"
		Me.ClientSize = New System.Drawing.Size(337, 377)
		Me.Location = New System.Drawing.Point(429, 214)
		Me.Font = New System.Drawing.Font("Arial Narrow", 18!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Icon = CType(resources.GetObject("frmSetData.Icon"), System.Drawing.Icon)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmSetData"
		Me.Frame6.Text = "Chroma"
		Me.Frame6.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame6.Size = New System.Drawing.Size(160, 54)
		Me.Frame6.Location = New System.Drawing.Point(170, 248)
		Me.Frame6.TabIndex = 30
		Me.Frame6.BackColor = System.Drawing.SystemColors.Control
		Me.Frame6.Enabled = True
		Me.Frame6.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame6.Visible = True
		Me.Frame6.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame6.Name = "Frame6"
		Me.cmbChromaModel.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbChromaModel.Size = New System.Drawing.Size(67, 23)
		Me.cmbChromaModel.Location = New System.Drawing.Point(80, 20)
		Me.cmbChromaModel.Items.AddRange(New Object(){"2401", "2402", "22293", "22293_A", "22293_B", "2233", "2233_A", "2233_B", "2333_B", "23293_B", "2234", "22294", "22294_A", "23294"})
		Me.cmbChromaModel.TabIndex = 32
		Me.cmbChromaModel.Text = "22294"
		Me.cmbChromaModel.BackColor = System.Drawing.SystemColors.Window
		Me.cmbChromaModel.CausesValidation = True
		Me.cmbChromaModel.Enabled = True
		Me.cmbChromaModel.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbChromaModel.IntegralHeight = True
		Me.cmbChromaModel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbChromaModel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbChromaModel.Sorted = False
		Me.cmbChromaModel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbChromaModel.TabStop = True
		Me.cmbChromaModel.Visible = True
		Me.cmbChromaModel.Name = "cmbChromaModel"
		Me.lbChromaModel.Text = "Model:"
		Me.lbChromaModel.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lbChromaModel.Size = New System.Drawing.Size(60, 17)
		Me.lbChromaModel.Location = New System.Drawing.Point(14, 22)
		Me.lbChromaModel.TabIndex = 31
		Me.lbChromaModel.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lbChromaModel.BackColor = System.Drawing.SystemColors.Control
		Me.lbChromaModel.Enabled = True
		Me.lbChromaModel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lbChromaModel.Cursor = System.Windows.Forms.Cursors.Default
		Me.lbChromaModel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lbChromaModel.UseMnemonic = True
		Me.lbChromaModel.Visible = True
		Me.lbChromaModel.AutoSize = False
		Me.lbChromaModel.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lbChromaModel.Name = "lbChromaModel"
		Me.Frame5.Text = "Serial Port"
		Me.Frame5.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame5.Size = New System.Drawing.Size(160, 73)
		Me.Frame5.Location = New System.Drawing.Point(170, 168)
		Me.Frame5.TabIndex = 25
		Me.Frame5.BackColor = System.Drawing.SystemColors.Control
		Me.Frame5.Enabled = True
		Me.Frame5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame5.Visible = True
		Me.Frame5.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame5.Name = "Frame5"
		Me.cmbComID.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbComID.Size = New System.Drawing.Size(67, 23)
		Me.cmbComID.Location = New System.Drawing.Point(80, 20)
		Me.cmbComID.TabIndex = 27
		Me.cmbComID.Text = "COM1"
		Me.cmbComID.BackColor = System.Drawing.SystemColors.Window
		Me.cmbComID.CausesValidation = True
		Me.cmbComID.Enabled = True
		Me.cmbComID.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbComID.IntegralHeight = True
		Me.cmbComID.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbComID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbComID.Sorted = False
		Me.cmbComID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbComID.TabStop = True
		Me.cmbComID.Visible = True
		Me.cmbComID.Name = "cmbComID"
		Me.cmbComBaud.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbComBaud.Size = New System.Drawing.Size(67, 23)
		Me.cmbComBaud.Location = New System.Drawing.Point(80, 44)
		Me.cmbComBaud.TabIndex = 26
		Me.cmbComBaud.Text = "9600"
		Me.cmbComBaud.BackColor = System.Drawing.SystemColors.Window
		Me.cmbComBaud.CausesValidation = True
		Me.cmbComBaud.Enabled = True
		Me.cmbComBaud.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbComBaud.IntegralHeight = True
		Me.cmbComBaud.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbComBaud.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbComBaud.Sorted = False
		Me.cmbComBaud.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbComBaud.TabStop = True
		Me.cmbComBaud.Visible = True
		Me.cmbComBaud.Name = "cmbComBaud"
		Me.lbComId.Text = "ComID:"
		Me.lbComId.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lbComId.Size = New System.Drawing.Size(60, 17)
		Me.lbComId.Location = New System.Drawing.Point(14, 22)
		Me.lbComId.TabIndex = 29
		Me.lbComId.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lbComId.BackColor = System.Drawing.SystemColors.Control
		Me.lbComId.Enabled = True
		Me.lbComId.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lbComId.Cursor = System.Windows.Forms.Cursors.Default
		Me.lbComId.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lbComId.UseMnemonic = True
		Me.lbComId.Visible = True
		Me.lbComId.AutoSize = False
		Me.lbComId.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lbComId.Name = "lbComId"
		Me.lbComBaud.Text = "ComBaud:"
		Me.lbComBaud.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lbComBaud.Size = New System.Drawing.Size(60, 17)
		Me.lbComBaud.Location = New System.Drawing.Point(14, 47)
		Me.lbComBaud.TabIndex = 28
		Me.lbComBaud.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lbComBaud.BackColor = System.Drawing.SystemColors.Control
		Me.lbComBaud.Enabled = True
		Me.lbComBaud.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lbComBaud.Cursor = System.Windows.Forms.Cursors.Default
		Me.lbComBaud.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lbComBaud.UseMnemonic = True
		Me.lbComBaud.Visible = True
		Me.lbComBaud.AutoSize = False
		Me.lbComBaud.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lbComBaud.Name = "lbComBaud"
		Me.Frame4.Text = "Communication Mode"
		Me.Frame4.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame4.Size = New System.Drawing.Size(160, 53)
		Me.Frame4.Location = New System.Drawing.Point(170, 112)
		Me.Frame4.TabIndex = 22
		Me.Frame4.BackColor = System.Drawing.SystemColors.Control
		Me.Frame4.Enabled = True
		Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame4.Visible = True
		Me.Frame4.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame4.Name = "Frame4"
		Me.optUart.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optUart.Text = "UART"
		Me.optUart.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optUart.Size = New System.Drawing.Size(54, 17)
		Me.optUart.Location = New System.Drawing.Point(8, 24)
		Me.optUart.TabIndex = 24
		Me.optUart.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optUart.BackColor = System.Drawing.SystemColors.Control
		Me.optUart.CausesValidation = True
		Me.optUart.Enabled = True
		Me.optUart.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optUart.Cursor = System.Windows.Forms.Cursors.Default
		Me.optUart.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optUart.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optUart.TabStop = True
		Me.optUart.Checked = False
		Me.optUart.Visible = True
		Me.optUart.Name = "optUart"
		Me.optNetwork.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optNetwork.Text = "Network"
		Me.optNetwork.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optNetwork.Size = New System.Drawing.Size(67, 17)
		Me.optNetwork.Location = New System.Drawing.Point(74, 24)
		Me.optNetwork.TabIndex = 23
		Me.optNetwork.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optNetwork.BackColor = System.Drawing.SystemColors.Control
		Me.optNetwork.CausesValidation = True
		Me.optNetwork.Enabled = True
		Me.optNetwork.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optNetwork.Cursor = System.Windows.Forms.Cursors.Default
		Me.optNetwork.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optNetwork.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optNetwork.TabStop = True
		Me.optNetwork.Checked = False
		Me.optNetwork.Visible = True
		Me.optNetwork.Name = "optNetwork"
		Me.Frame2.Text = "Common Setting"
		Me.Frame2.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.Size = New System.Drawing.Size(160, 121)
		Me.Frame2.Location = New System.Drawing.Point(8, 248)
		Me.Frame2.TabIndex = 10
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame2.Name = "Frame2"
		Me.txtLvSpec.AutoSize = False
		Me.txtLvSpec.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtLvSpec.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLvSpec.Size = New System.Drawing.Size(67, 20)
		Me.txtLvSpec.Location = New System.Drawing.Point(80, 90)
		Me.txtLvSpec.TabIndex = 21
		Me.txtLvSpec.Text = "280"
		Me.txtLvSpec.AcceptsReturn = True
		Me.txtLvSpec.BackColor = System.Drawing.SystemColors.Window
		Me.txtLvSpec.CausesValidation = True
		Me.txtLvSpec.Enabled = True
		Me.txtLvSpec.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLvSpec.HideSelection = True
		Me.txtLvSpec.ReadOnly = False
		Me.txtLvSpec.Maxlength = 0
		Me.txtLvSpec.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLvSpec.MultiLine = False
		Me.txtLvSpec.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLvSpec.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLvSpec.TabStop = True
		Me.txtLvSpec.Visible = True
		Me.txtLvSpec.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtLvSpec.Name = "txtLvSpec"
		Me.cmbInputSource.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbInputSource.Size = New System.Drawing.Size(67, 23)
		Me.cmbInputSource.Location = New System.Drawing.Point(80, 67)
		Me.cmbInputSource.Items.AddRange(New Object(){"HDMI1", "HDMI2", "HDMI3", "AV1"})
		Me.cmbInputSource.TabIndex = 17
		Me.cmbInputSource.Text = "HDMI1"
		Me.cmbInputSource.BackColor = System.Drawing.SystemColors.Window
		Me.cmbInputSource.CausesValidation = True
		Me.cmbInputSource.Enabled = True
		Me.cmbInputSource.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbInputSource.IntegralHeight = True
		Me.cmbInputSource.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbInputSource.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbInputSource.Sorted = False
		Me.cmbInputSource.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbInputSource.TabStop = True
		Me.cmbInputSource.Visible = True
		Me.cmbInputSource.Name = "cmbInputSource"
		Me.txtSNLen.AutoSize = False
		Me.txtSNLen.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtSNLen.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSNLen.Size = New System.Drawing.Size(67, 20)
		Me.txtSNLen.Location = New System.Drawing.Point(80, 44)
		Me.txtSNLen.TabIndex = 13
		Me.txtSNLen.Text = "1"
		Me.txtSNLen.AcceptsReturn = True
		Me.txtSNLen.BackColor = System.Drawing.SystemColors.Window
		Me.txtSNLen.CausesValidation = True
		Me.txtSNLen.Enabled = True
		Me.txtSNLen.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSNLen.HideSelection = True
		Me.txtSNLen.ReadOnly = False
		Me.txtSNLen.Maxlength = 0
		Me.txtSNLen.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSNLen.MultiLine = False
		Me.txtSNLen.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSNLen.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSNLen.TabStop = True
		Me.txtSNLen.Visible = True
		Me.txtSNLen.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtSNLen.Name = "txtSNLen"
		Me.txtDelay.AutoSize = False
		Me.txtDelay.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtDelay.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDelay.Size = New System.Drawing.Size(67, 20)
		Me.txtDelay.Location = New System.Drawing.Point(80, 20)
		Me.txtDelay.TabIndex = 11
		Me.txtDelay.Text = "500"
		Me.txtDelay.AcceptsReturn = True
		Me.txtDelay.BackColor = System.Drawing.SystemColors.Window
		Me.txtDelay.CausesValidation = True
		Me.txtDelay.Enabled = True
		Me.txtDelay.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDelay.HideSelection = True
		Me.txtDelay.ReadOnly = False
		Me.txtDelay.Maxlength = 0
		Me.txtDelay.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDelay.MultiLine = False
		Me.txtDelay.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDelay.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDelay.TabStop = True
		Me.txtDelay.Visible = True
		Me.txtDelay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtDelay.Name = "txtDelay"
		Me.lbLvSpec.Text = "Lv Spec:"
		Me.lbLvSpec.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lbLvSpec.Size = New System.Drawing.Size(60, 17)
		Me.lbLvSpec.Location = New System.Drawing.Point(14, 92)
		Me.lbLvSpec.TabIndex = 20
		Me.lbLvSpec.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lbLvSpec.BackColor = System.Drawing.SystemColors.Control
		Me.lbLvSpec.Enabled = True
		Me.lbLvSpec.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lbLvSpec.Cursor = System.Windows.Forms.Cursors.Default
		Me.lbLvSpec.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lbLvSpec.UseMnemonic = True
		Me.lbLvSpec.Visible = True
		Me.lbLvSpec.AutoSize = False
		Me.lbLvSpec.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lbLvSpec.Name = "lbLvSpec"
		Me.lbInputSrc.Text = "TV Source:"
		Me.lbInputSrc.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lbInputSrc.Size = New System.Drawing.Size(60, 17)
		Me.lbInputSrc.Location = New System.Drawing.Point(14, 69)
		Me.lbInputSrc.TabIndex = 18
		Me.lbInputSrc.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lbInputSrc.BackColor = System.Drawing.SystemColors.Control
		Me.lbInputSrc.Enabled = True
		Me.lbInputSrc.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lbInputSrc.Cursor = System.Windows.Forms.Cursors.Default
		Me.lbInputSrc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lbInputSrc.UseMnemonic = True
		Me.lbInputSrc.Visible = True
		Me.lbInputSrc.AutoSize = False
		Me.lbInputSrc.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lbInputSrc.Name = "lbInputSrc"
		Me.lbSnLen.Text = "SN_Len:"
		Me.lbSnLen.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lbSnLen.Size = New System.Drawing.Size(60, 17)
		Me.lbSnLen.Location = New System.Drawing.Point(14, 46)
		Me.lbSnLen.TabIndex = 14
		Me.lbSnLen.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lbSnLen.BackColor = System.Drawing.SystemColors.Control
		Me.lbSnLen.Enabled = True
		Me.lbSnLen.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lbSnLen.Cursor = System.Windows.Forms.Cursors.Default
		Me.lbSnLen.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lbSnLen.UseMnemonic = True
		Me.lbSnLen.Visible = True
		Me.lbSnLen.AutoSize = False
		Me.lbSnLen.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lbSnLen.Name = "lbSnLen"
		Me.lbDelayMs.Text = "Delay(ms):"
		Me.lbDelayMs.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lbDelayMs.Size = New System.Drawing.Size(60, 17)
		Me.lbDelayMs.Location = New System.Drawing.Point(14, 22)
		Me.lbDelayMs.TabIndex = 12
		Me.lbDelayMs.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lbDelayMs.BackColor = System.Drawing.SystemColors.Control
		Me.lbDelayMs.Enabled = True
		Me.lbDelayMs.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lbDelayMs.Cursor = System.Windows.Forms.Cursors.Default
		Me.lbDelayMs.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lbDelayMs.UseMnemonic = True
		Me.lbDelayMs.Visible = True
		Me.lbDelayMs.AutoSize = False
		Me.lbDelayMs.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lbDelayMs.Name = "lbDelayMs"
		Me.Frame3.Text = "CA310/CA210"
		Me.Frame3.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame3.Size = New System.Drawing.Size(160, 54)
		Me.Frame3.Location = New System.Drawing.Point(170, 56)
		Me.Frame3.TabIndex = 7
		Me.Frame3.BackColor = System.Drawing.SystemColors.Control
		Me.Frame3.Enabled = True
		Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame3.Visible = True
		Me.Frame3.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame3.Name = "Frame3"
		Me.txtChannel.AutoSize = False
		Me.txtChannel.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtChannel.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtChannel.Size = New System.Drawing.Size(67, 20)
		Me.txtChannel.Location = New System.Drawing.Point(80, 20)
		Me.txtChannel.TabIndex = 9
		Me.txtChannel.Text = "1"
		Me.txtChannel.AcceptsReturn = True
		Me.txtChannel.BackColor = System.Drawing.SystemColors.Window
		Me.txtChannel.CausesValidation = True
		Me.txtChannel.Enabled = True
		Me.txtChannel.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtChannel.HideSelection = True
		Me.txtChannel.ReadOnly = False
		Me.txtChannel.Maxlength = 0
		Me.txtChannel.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtChannel.MultiLine = False
		Me.txtChannel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtChannel.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtChannel.TabStop = True
		Me.txtChannel.Visible = True
		Me.txtChannel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.txtChannel.Name = "txtChannel"
		Me.lbChannelNum.Text = "Channel:"
		Me.lbChannelNum.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lbChannelNum.Size = New System.Drawing.Size(60, 17)
		Me.lbChannelNum.Location = New System.Drawing.Point(14, 22)
		Me.lbChannelNum.TabIndex = 8
		Me.lbChannelNum.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lbChannelNum.BackColor = System.Drawing.SystemColors.Control
		Me.lbChannelNum.Enabled = True
		Me.lbChannelNum.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lbChannelNum.Cursor = System.Windows.Forms.Cursors.Default
		Me.lbChannelNum.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lbChannelNum.UseMnemonic = True
		Me.lbChannelNum.Visible = True
		Me.lbChannelNum.AutoSize = False
		Me.lbChannelNum.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lbChannelNum.Name = "lbChannelNum"
		Me.Frame1.Text = "Selection"
		Me.Frame1.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.Size = New System.Drawing.Size(160, 189)
		Me.Frame1.Location = New System.Drawing.Point(8, 56)
		Me.Frame1.TabIndex = 6
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.Check7.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.Check7.Text = "Adjust Offset"
		Me.Check7.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Check7.Size = New System.Drawing.Size(127, 17)
		Me.Check7.Location = New System.Drawing.Point(14, 164)
		Me.Check7.TabIndex = 16
		Me.Check7.CheckState = System.Windows.Forms.CheckState.Checked
		Me.Check7.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.Check7.BackColor = System.Drawing.SystemColors.Control
		Me.Check7.CausesValidation = True
		Me.Check7.Enabled = True
		Me.Check7.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Check7.Cursor = System.Windows.Forms.Cursors.Default
		Me.Check7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Check7.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Check7.TabStop = True
		Me.Check7.Visible = True
		Me.Check7.Name = "Check7"
		Me.Check6.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.Check6.Text = "Check Color"
		Me.Check6.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Check6.Size = New System.Drawing.Size(127, 17)
		Me.Check6.Location = New System.Drawing.Point(14, 140)
		Me.Check6.TabIndex = 15
		Me.Check6.CheckState = System.Windows.Forms.CheckState.Checked
		Me.Check6.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.Check6.BackColor = System.Drawing.SystemColors.Control
		Me.Check6.CausesValidation = True
		Me.Check6.Enabled = True
		Me.Check6.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Check6.Cursor = System.Windows.Forms.Cursors.Default
		Me.Check6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Check6.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Check6.TabStop = True
		Me.Check6.Visible = True
		Me.Check6.Name = "Check6"
		Me.Check1.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.Check1.Text = "COOL_2"
		Me.Check1.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Check1.Size = New System.Drawing.Size(127, 17)
		Me.Check1.Location = New System.Drawing.Point(14, 24)
		Me.Check1.TabIndex = 0
		Me.Check1.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.Check1.BackColor = System.Drawing.SystemColors.Control
		Me.Check1.CausesValidation = True
		Me.Check1.Enabled = True
		Me.Check1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Check1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Check1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Check1.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Check1.TabStop = True
		Me.Check1.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.Check1.Visible = True
		Me.Check1.Name = "Check1"
		Me.Check2.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.Check2.Text = "COOL_1"
		Me.Check2.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Check2.Size = New System.Drawing.Size(127, 17)
		Me.Check2.Location = New System.Drawing.Point(14, 47)
		Me.Check2.TabIndex = 1
		Me.Check2.CheckState = System.Windows.Forms.CheckState.Checked
		Me.Check2.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.Check2.BackColor = System.Drawing.SystemColors.Control
		Me.Check2.CausesValidation = True
		Me.Check2.Enabled = True
		Me.Check2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Check2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Check2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Check2.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Check2.TabStop = True
		Me.Check2.Visible = True
		Me.Check2.Name = "Check2"
		Me.Check3.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.Check3.Text = "NORMAL"
		Me.Check3.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Check3.Size = New System.Drawing.Size(127, 17)
		Me.Check3.Location = New System.Drawing.Point(14, 70)
		Me.Check3.TabIndex = 2
		Me.Check3.CheckState = System.Windows.Forms.CheckState.Checked
		Me.Check3.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.Check3.BackColor = System.Drawing.SystemColors.Control
		Me.Check3.CausesValidation = True
		Me.Check3.Enabled = True
		Me.Check3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Check3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Check3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Check3.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Check3.TabStop = True
		Me.Check3.Visible = True
		Me.Check3.Name = "Check3"
		Me.Check4.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.Check4.Text = "WARM_1"
		Me.Check4.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Check4.Size = New System.Drawing.Size(127, 17)
		Me.Check4.Location = New System.Drawing.Point(14, 94)
		Me.Check4.TabIndex = 3
		Me.Check4.CheckState = System.Windows.Forms.CheckState.Checked
		Me.Check4.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.Check4.BackColor = System.Drawing.SystemColors.Control
		Me.Check4.CausesValidation = True
		Me.Check4.Enabled = True
		Me.Check4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Check4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Check4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Check4.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Check4.TabStop = True
		Me.Check4.Visible = True
		Me.Check4.Name = "Check4"
		Me.Check5.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.Check5.Text = "WARM_2"
		Me.Check5.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Check5.Size = New System.Drawing.Size(127, 17)
		Me.Check5.Location = New System.Drawing.Point(14, 117)
		Me.Check5.TabIndex = 4
		Me.Check5.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.Check5.BackColor = System.Drawing.SystemColors.Control
		Me.Check5.CausesValidation = True
		Me.Check5.Enabled = True
		Me.Check5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Check5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Check5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Check5.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Check5.TabStop = True
		Me.Check5.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.Check5.Visible = True
		Me.Check5.Name = "Check5"
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command1.Text = "Save"
		Me.Command1.Font = New System.Drawing.Font("Arial", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command1.Size = New System.Drawing.Size(73, 29)
		Me.Command1.Location = New System.Drawing.Point(256, 336)
		Me.Command1.TabIndex = 5
		Me.Command1.BackColor = System.Drawing.SystemColors.Control
		Me.Command1.CausesValidation = True
		Me.Command1.Enabled = True
		Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command1.TabStop = True
		Me.Command1.Name = "Command1"
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label1.Text = "Label1"
		Me.Label1.Font = New System.Drawing.Font("Arial", 24!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(321, 37)
		Me.Label1.Location = New System.Drawing.Point(8, 8)
		Me.Label1.TabIndex = 19
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(Frame6)
		Me.Controls.Add(Frame5)
		Me.Controls.Add(Frame4)
		Me.Controls.Add(Frame2)
		Me.Controls.Add(Frame3)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(Command1)
		Me.Controls.Add(Label1)
		Me.Frame6.Controls.Add(cmbChromaModel)
		Me.Frame6.Controls.Add(lbChromaModel)
		Me.Frame5.Controls.Add(cmbComID)
		Me.Frame5.Controls.Add(cmbComBaud)
		Me.Frame5.Controls.Add(lbComId)
		Me.Frame5.Controls.Add(lbComBaud)
		Me.Frame4.Controls.Add(optUart)
		Me.Frame4.Controls.Add(optNetwork)
		Me.Frame2.Controls.Add(txtLvSpec)
		Me.Frame2.Controls.Add(cmbInputSource)
		Me.Frame2.Controls.Add(txtSNLen)
		Me.Frame2.Controls.Add(txtDelay)
		Me.Frame2.Controls.Add(lbLvSpec)
		Me.Frame2.Controls.Add(lbInputSrc)
		Me.Frame2.Controls.Add(lbSnLen)
		Me.Frame2.Controls.Add(lbDelayMs)
		Me.Frame3.Controls.Add(txtChannel)
		Me.Frame3.Controls.Add(lbChannelNum)
		Me.Frame1.Controls.Add(Check7)
		Me.Frame1.Controls.Add(Check6)
		Me.Frame1.Controls.Add(Check1)
		Me.Frame1.Controls.Add(Check2)
		Me.Frame1.Controls.Add(Check3)
		Me.Frame1.Controls.Add(Check4)
		Me.Frame1.Controls.Add(Check5)
		Me.Frame6.ResumeLayout(False)
		Me.Frame5.ResumeLayout(False)
		Me.Frame4.ResumeLayout(False)
		Me.Frame2.ResumeLayout(False)
		Me.Frame3.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class