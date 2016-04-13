<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSplash
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
	Public WithEvents cmbModelName As System.Windows.Forms.ComboBox
	Public WithEvents lblVersion As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSplash))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmbModelName = New System.Windows.Forms.ComboBox
		Me.lblVersion = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ControlBox = False
		Me.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.ClientSize = New System.Drawing.Size(372, 185)
		Me.Location = New System.Drawing.Point(17, 94)
		Me.KeyPreview = True
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.Enabled = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmSplash"
		Me.cmbModelName.BackColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me.cmbModelName.Font = New System.Drawing.Font("Arial", 20.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbModelName.ForeColor = System.Drawing.Color.Blue
		Me.cmbModelName.Size = New System.Drawing.Size(209, 40)
		Me.cmbModelName.Location = New System.Drawing.Point(72, 96)
		Me.cmbModelName.Sorted = True
		Me.cmbModelName.TabIndex = 0
		Me.cmbModelName.Text = "Sample1"
		Me.cmbModelName.CausesValidation = True
		Me.cmbModelName.Enabled = True
		Me.cmbModelName.IntegralHeight = True
		Me.cmbModelName.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbModelName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbModelName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbModelName.TabStop = True
		Me.cmbModelName.Visible = True
		Me.cmbModelName.Name = "cmbModelName"
		Me.lblVersion.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblVersion.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.lblVersion.Text = "Version "
		Me.lblVersion.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblVersion.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.lblVersion.Size = New System.Drawing.Size(55, 17)
		Me.lblVersion.Location = New System.Drawing.Point(312, 48)
		Me.lblVersion.TabIndex = 4
		Me.lblVersion.Enabled = True
		Me.lblVersion.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblVersion.UseMnemonic = True
		Me.lblVersion.Visible = True
		Me.lblVersion.AutoSize = True
		Me.lblVersion.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblVersion.Name = "lblVersion"
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label2.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.Label2.Text = "Auto White Balance System"
		Me.Label2.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.ForeColor = System.Drawing.Color.Black
		Me.Label2.Size = New System.Drawing.Size(357, 22)
		Me.Label2.Location = New System.Drawing.Point(8, 16)
		Me.Label2.TabIndex = 3
		Me.Label2.Enabled = True
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = True
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.Label1.Text = "Please select your model:"
		Me.Label1.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.ForeColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.Label1.Size = New System.Drawing.Size(217, 25)
		Me.Label1.Location = New System.Drawing.Point(64, 72)
		Me.Label1.TabIndex = 2
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.Enabled = True
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label3.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.Label3.Text = "Copyright 2013-2016    Design by ECHOM"
		Me.Label3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.Label3.Size = New System.Drawing.Size(353, 17)
		Me.Label3.Location = New System.Drawing.Point(8, 160)
		Me.Label3.TabIndex = 1
		Me.Label3.Enabled = True
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Controls.Add(cmbModelName)
		Me.Controls.Add(lblVersion)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.Controls.Add(Label3)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class