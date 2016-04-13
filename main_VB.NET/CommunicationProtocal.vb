Option Strict Off
Option Explicit On
Interface _CommunicationProtocal
	Sub EnterFacMode()
	Sub ExitFacMode()
	Sub SwitchInputSource(ByRef strInputSrc As String, ByRef intSrcNum As Short)
	Sub ResetPicMode()
	Sub SetBrightness(ByRef intBrightness As Short)
	Sub SetContrast(ByRef intContrast As Short)
	Sub SetBacklight(ByRef intBacklight As Short)
	Sub SelColorTemp(ByRef strColorT As String, ByRef strInputSrc As String, ByRef intSrcNum As Short)
	Sub SetRGBGain(ByRef lngRGain As Integer, ByRef lngGGain As Integer, ByRef lngBGain As Integer)
	Sub SetRGBOffset(ByRef lngROffset As Integer, ByRef lngGOffset As Integer, ByRef lngBOffset As Integer)
	Sub SaveWBDataToAllSrc(ByRef strInputSrc As String, ByRef intSrcNum As Short)
End Interface
Friend Class CommunicationProtocal
	Implements _CommunicationProtocal
	'**********************************************
	' Interface for handling protocal.
	'**********************************************
	
	
	Public Sub EnterFacMode() Implements _CommunicationProtocal.EnterFacMode
		
	End Sub
	
	Public Sub ExitFacMode() Implements _CommunicationProtocal.ExitFacMode
		
	End Sub
	
	Public Sub SwitchInputSource(ByRef strInputSrc As String, ByRef intSrcNum As Short) Implements _CommunicationProtocal.SwitchInputSource
		
	End Sub
	
	Public Sub ResetPicMode() Implements _CommunicationProtocal.ResetPicMode
		
	End Sub
	
	Public Sub SetBrightness(ByRef intBrightness As Short) Implements _CommunicationProtocal.SetBrightness
		
	End Sub
	
	Public Sub SetContrast(ByRef intContrast As Short) Implements _CommunicationProtocal.SetContrast
		
	End Sub
	
	Public Sub SetBacklight(ByRef intBacklight As Short) Implements _CommunicationProtocal.SetBacklight
		
	End Sub
	
	Public Sub SelColorTemp(ByRef strColorT As String, ByRef strInputSrc As String, ByRef intSrcNum As Short) Implements _CommunicationProtocal.SelColorTemp
		
	End Sub
	
	Public Sub SetRGBGain(ByRef lngRGain As Integer, ByRef lngGGain As Integer, ByRef lngBGain As Integer) Implements _CommunicationProtocal.SetRGBGain
		
	End Sub
	
	Public Sub SetRGBOffset(ByRef lngROffset As Integer, ByRef lngGOffset As Integer, ByRef lngBOffset As Integer) Implements _CommunicationProtocal.SetRGBOffset
		
	End Sub
	
	Public Sub SaveWBDataToAllSrc(ByRef strInputSrc As String, ByRef intSrcNum As Short) Implements _CommunicationProtocal.SaveWBDataToAllSrc
		
	End Sub
End Class