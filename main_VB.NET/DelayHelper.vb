Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module DelayHelper
	
	Public Sub Delay(ByRef mmSec As Integer)
		On Error GoTo ShowError
		Dim start As Single
		start = VB.Timer()
		While (VB.Timer() - start) < (mmSec / 1000#)
			System.Windows.Forms.Application.DoEvents()
			
			If IsStop = True Then
				Exit Sub
			End If
		End While
		Exit Sub
		
ShowError: 
		MsgBox(Err.Source & "------" & Err.Description)
		Exit Sub
	End Sub
	
	
	Public Sub DelayMS(ByRef mmSec As Integer)
		On Error GoTo ShowError
		Dim start As Single
		start = VB.Timer()
		While (VB.Timer() - start) < (mmSec / 1000#)
			System.Windows.Forms.Application.DoEvents()
			
			If IsStop = True Then
				Exit Sub
			End If
			
		End While
		Exit Sub
		
ShowError: 
		MsgBox(Err.Source & "------" & Err.Description)
		Exit Sub
	End Sub
	
	Public Sub DelaySWithFlag(ByRef Sec As Integer, ByRef flag As Boolean)
		On Error GoTo ShowError
		Dim start As Single
		start = VB.Timer()
		While (VB.Timer() - start) < Sec
			System.Windows.Forms.Application.DoEvents()
			
			If flag = True Then
				Exit Sub
			End If
			
			If IsStop = True Then
				Exit Sub
			End If
		End While
		Exit Sub
		
ShowError: 
		MsgBox(Err.Source & "------" & Err.Description)
		Exit Sub
	End Sub
End Module