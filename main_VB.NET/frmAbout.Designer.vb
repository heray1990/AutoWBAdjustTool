<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAbout
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
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents lblDisclaimer As System.Windows.Forms.Label
	Public WithEvents lblDescription As System.Windows.Forms.Label
	Public WithEvents _Line1_1 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents lblTitle As System.Windows.Forms.Label
	Public WithEvents _Line1_0 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents lblVersion As System.Windows.Forms.Label
	Public WithEvents Line1 As LineShapeArray
	Public WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAbout))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.cmdOK = New System.Windows.Forms.Button
		Me.lblDisclaimer = New System.Windows.Forms.Label
		Me.lblDescription = New System.Windows.Forms.Label
		Me._Line1_1 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.lblTitle = New System.Windows.Forms.Label
		Me._Line1_0 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.lblVersion = New System.Windows.Forms.Label
		Me.Line1 = New LineShapeArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Line1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "About System"
		Me.ClientSize = New System.Drawing.Size(382, 175)
		Me.Location = New System.Drawing.Point(156, 129)
		Me.Icon = CType(resources.GetObject("frmAbout.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmAbout"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdOK
		Me.cmdOK.Text = "OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(84, 23)
		Me.cmdOK.Location = New System.Drawing.Point(288, 136)
		Me.cmdOK.TabIndex = 0
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.lblDisclaimer.Text = "Copyright 2013 - 2016             All rights reserved.              Echom "
		Me.lblDisclaimer.ForeColor = System.Drawing.Color.Black
		Me.lblDisclaimer.Size = New System.Drawing.Size(201, 37)
		Me.lblDisclaimer.Location = New System.Drawing.Point(16, 128)
		Me.lblDisclaimer.TabIndex = 4
		Me.lblDisclaimer.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDisclaimer.BackColor = System.Drawing.SystemColors.Control
		Me.lblDisclaimer.Enabled = True
		Me.lblDisclaimer.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDisclaimer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDisclaimer.UseMnemonic = True
		Me.lblDisclaimer.Visible = True
		Me.lblDisclaimer.AutoSize = False
		Me.lblDisclaimer.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDisclaimer.Name = "lblDisclaimer"
		Me.lblDescription.Text = "Auto White Balance adjustment by Echom."
		Me.lblDescription.Font = New System.Drawing.Font("PMingLiU", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
		Me.lblDescription.ForeColor = System.Drawing.Color.Black
		Me.lblDescription.Size = New System.Drawing.Size(301, 46)
		Me.lblDescription.Location = New System.Drawing.Point(72, 72)
		Me.lblDescription.TabIndex = 3
		Me.lblDescription.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDescription.BackColor = System.Drawing.SystemColors.Control
		Me.lblDescription.Enabled = True
		Me.lblDescription.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDescription.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDescription.UseMnemonic = True
		Me.lblDescription.Visible = True
		Me.lblDescription.AutoSize = False
		Me.lblDescription.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDescription.Name = "lblDescription"
		Me._Line1_1.BorderColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me._Line1_1.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line1_1.X1 = 8
		Me._Line1_1.X2 = 356
		Me._Line1_1.Y1 = 83
		Me._Line1_1.Y2 = 83
		Me._Line1_1.BorderWidth = 1
		Me._Line1_1.Visible = True
		Me._Line1_1.Name = "_Line1_1"
		Me.lblTitle.Text = "Auto White Balance System"
		Me.lblTitle.Font = New System.Drawing.Font("PMingLiU", 14.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
		Me.lblTitle.ForeColor = System.Drawing.Color.Black
		Me.lblTitle.Size = New System.Drawing.Size(299, 32)
		Me.lblTitle.Location = New System.Drawing.Point(72, 16)
		Me.lblTitle.TabIndex = 1
		Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblTitle.BackColor = System.Drawing.SystemColors.Control
		Me.lblTitle.Enabled = True
		Me.lblTitle.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTitle.UseMnemonic = True
		Me.lblTitle.Visible = True
		Me.lblTitle.AutoSize = False
		Me.lblTitle.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblTitle.Name = "lblTitle"
		Me._Line1_0.BorderColor = System.Drawing.Color.White
		Me._Line1_0.BorderWidth = 2
		Me._Line1_0.X1 = 8
		Me._Line1_0.X2 = 355
		Me._Line1_0.Y1 = 83
		Me._Line1_0.Y2 = 83
		Me._Line1_0.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line1_0.Visible = True
		Me._Line1_0.Name = "_Line1_0"
		Me.lblVersion.Text = "Ver1.00"
		Me.lblVersion.Size = New System.Drawing.Size(259, 15)
		Me.lblVersion.Location = New System.Drawing.Point(72, 52)
		Me.lblVersion.TabIndex = 2
		Me.lblVersion.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblVersion.BackColor = System.Drawing.SystemColors.Control
		Me.lblVersion.Enabled = True
		Me.lblVersion.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblVersion.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblVersion.UseMnemonic = True
		Me.lblVersion.Visible = True
		Me.lblVersion.AutoSize = False
		Me.lblVersion.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblVersion.Name = "lblVersion"
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(lblDisclaimer)
		Me.Controls.Add(lblDescription)
		Me.ShapeContainer1.Shapes.Add(_Line1_1)
		Me.Controls.Add(lblTitle)
		Me.ShapeContainer1.Shapes.Add(_Line1_0)
		Me.Controls.Add(lblVersion)
		Me.Controls.Add(ShapeContainer1)
		Me.Line1.SetIndex(_Line1_1, CType(1, Short))
		Me.Line1.SetIndex(_Line1_0, CType(0, Short))
		CType(Me.Line1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class