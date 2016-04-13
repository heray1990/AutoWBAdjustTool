Option Strict Off
Option Explicit On
Friend Class frmSplash
	Inherits System.Windows.Forms.Form
	
	
	
	Private Sub frmSplash_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Click
		Me.Close()
	End Sub
	
	Private Sub frmSplash_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.DoubleClick
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Form event frmSplash.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmSplash_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
		Me.Close()
	End Sub
	
	Private Sub frmSplash_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Me.Close()
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub frmSplash_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		On Error GoTo ErrExit
		Dim strProjectName As Object
		
		cmbModelName.Items.Clear()
		
		For	Each strProjectName In GetProjectList
			'UPGRADE_WARNING: Couldn't resolve default property of object strProjectName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmbModelName.Items.Add(strProjectName)
		Next strProjectName
		
		lblVersion.Text = "Version " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision
		
		cmbModelName.Text = GetCurProjectName
		
		Exit Sub
		
ErrExit: 
		MsgBox(Err.Description, MsgBoxStyle.Critical, Err.Source)
	End Sub
	
	Private Sub frmSplash_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		
		'On Error GoTo ErrExit
		gstrCurProjName = cmbModelName.Text
		SetCurProjectName(gstrCurProjName)
		
		Form1.Show()
		Exit Sub
		
		'ErrExit:
		'MsgBox ("The Licence Key is Wrong.")
	End Sub
End Class