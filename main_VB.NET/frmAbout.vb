Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmAbout
	Inherits System.Windows.Forms.Form
	
	
	Const READ_CONTROL As Integer = &H20000
	Const KEY_QUERY_VALUE As Integer = &H1
	Const KEY_SET_VALUE As Integer = &H2
	Const KEY_CREATE_SUB_KEY As Integer = &H4
	Const KEY_ENUMERATE_SUB_KEYS As Integer = &H8
	Const KEY_NOTIFY As Integer = &H10
	Const KEY_CREATE_LINK As Integer = &H20
	Const KEY_ALL_ACCESS As Double = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
	
	
	Const HKEY_LOCAL_MACHINE As Integer = &H80000002
	Const ERROR_SUCCESS As Short = 0
	Const REG_SZ As Short = 1
	Const REG_DWORD As Short = 4
	
	Const gREGKEYSYSINFOLOC As String = "SOFTWARE\Microsoft\Shared Tools Location"
	Const gREGVALSYSINFOLOC As String = "MSINFO"
	Const gREGKEYSYSINFO As String = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
	Const gREGVALSYSINFO As String = "PATH"
	
	Private Declare Function RegOpenKeyEx Lib "advapi32"  Alias "RegOpenKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
	Private Declare Function RegQueryValueEx Lib "advapi32"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As String, ByRef lpcbData As Integer) As Integer
	Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Integer) As Integer
	
	
	Private Sub cmdSysInfo_Click()
		'Call StartSysInfo
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Me.Close()
		'UPGRADE_WARNING: Form method Form1.ZOrder has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		Form1.BringToFront()
	End Sub
	
	Private Sub frmAbout_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Text = "About " & My.Application.Info.Title
		lblVersion.Text = "Version " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision
		lblTitle.Text = My.Application.Info.Title
	End Sub
	
	Public Sub StartSysInfo()
		On Error GoTo SysInfoErr
		
		Dim rc As Integer
		Dim SysInfoPath As String
		
		
		If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
			
		ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then 
			
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
				SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
				
				
			Else
				GoTo SysInfoErr
			End If
			
		Else
			GoTo SysInfoErr
		End If
		
		Call Shell(SysInfoPath, AppWinStyle.NormalFocus)
		
		Exit Sub
SysInfoErr: 
		MsgBox("No system information now.", MsgBoxStyle.OKOnly)
	End Sub
	
	Public Function GetKeyValue(ByRef KeyRoot As Integer, ByRef KeyName As String, ByRef SubKeyRef As String, ByRef KeyVal As String) As Boolean
		Dim i As Integer
		Dim rc As Integer
		Dim hKey As Integer
		Dim hDepth As Integer
		Dim KeyValType As Integer
		Dim tmpVal As String
		Dim KeyValSize As Integer
		
		rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
		
		If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
		
		tmpVal = New String(Chr(0), 1024)
		KeyValSize = 1024
		
		
		rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)
		
		If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
		
		If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then
			tmpVal = VB.Left(tmpVal, KeyValSize - 1)
		Else
			tmpVal = VB.Left(tmpVal, KeyValSize)
		End If
		
		Select Case KeyValType
			Case REG_SZ
				KeyVal = tmpVal
			Case REG_DWORD
				For i = Len(tmpVal) To 1 Step -1
					KeyVal = KeyVal & Hex(Asc(Mid(tmpVal, i, 1)))
				Next 
				KeyVal = VB6.Format("&h" & KeyVal)
		End Select
		
		GetKeyValue = True
		rc = RegCloseKey(hKey)
		Exit Function
		
GetKeyError: 
		KeyVal = ""
		GetKeyValue = False
		rc = RegCloseKey(hKey)
	End Function
End Class