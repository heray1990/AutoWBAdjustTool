Option Strict Off
Option Explicit On
Module DeclareVariables
	
	Public Const FixR As Short = 1
	Public Const FixG As Short = 2
	Public Const FixB As Short = 3
	Public Const adjustMode1 As Short = 1
	Public Const adjustMode2 As Short = 2
	Public Const adjustMode3 As Short = 3
	Public Const adjustMode4 As Short = 4
	
	Public Const AdjustSingle As Short = 1
	Public Const AdjustDouble As Short = 0
	
	Public Const SingleStep As Short = 0
	Public Const ComplexStep As Short = 1
	
	Public Const HighBri As Short = 1
	Public Const LowBri As Short = 0
	
	Public Const cstrColorTempCool1 As String = "COOL1"
	Public Const cstrColorTempNormal As String = "NORMAL"
	Public Const cstrColorTempWarm1 As String = "WARM1"
	
	Public Const lastChkShwDataStep As Short = 6
	Public Const cmdReceiveWaitS As Short = 5
	Public Const strRemoteHost As String = "192.168.1.11"
	Public Const lngRemotePort As Integer = 8888
	
	Structure ColorTemp
		Dim x As Single
		Dim y As Single
		Dim lv As Single
	End Structure
	
	Public strBuff As String
	
	Public i As Short
	Public adjustGainAgainCool1Flag As Short
	Public adjustGainAgainNormalFlag As Short
	Public adjustGainAgainWarm1Flag As Short
	
	Public Const xxf As Short = 1
	Public Const xfyf As Short = 2
	Public Const yyf As Short = 3
	Public Const microStep As Boolean = True
	Public Const StepbyStep As Boolean = False
	
	Public Ca210ChannelNO As Integer
	Public delayTime As Integer
	
	Public isAdjustCool1 As Boolean
	Public isAdjustCool2 As Boolean
	Public isAdjustNormal As Boolean
	Public isAdjustWarm1 As Boolean
	Public isAdjustWarm2 As Boolean
	
	Public gintBarCodeLen As Short
	Public IsFunctionAutoBri As Boolean
	Public isCheckColorTemp As Boolean
	
	Public isAdjustOffset As Boolean
	
	Public gstrCurProjName As String
	Public IsStop As Boolean
	Public IsACK As Boolean
	Public setTVCurrentComID As Short
	Public setTVCurrentComBaud As Integer
	Public setTVInputSource As String
	Public setTVInputSourcePortNum As Short
	Public maxBrightnessSpec As Integer
	
	Public IsCa210ok As Boolean
	Public isUartMode As Boolean
	Public isNetworkConnected As Boolean
	
	Public gstrBarCode As String
	Public countTime As Integer
	Public gstrBrand As String
	
	Public gstrVPGModel As String
	
	
	Public Structure COLORTEMPSPEC
		Dim xx As Integer
		Dim yy As Integer
		Dim lv As Integer
		Dim nColorRR As Integer
		Dim nColorGG As Integer
		Dim nColorBB As Integer
		Dim xt As Integer
		Dim yt As Integer
		Dim nLowRR As Integer
		Dim nLowGG As Integer
		Dim nLowBB As Integer
	End Structure
	
	Public Structure REALCOLOR
		Dim xx As Integer
		Dim yy As Integer
		Dim lv As Integer
	End Structure
	
	Public Structure REALRGB
		Dim cRR As Integer
		Dim cGG As Integer
		Dim cBB As Integer
	End Structure
	Public rRGB As REALRGB
	Public rRGB1 As REALRGB
	
	Public Enum CommunicationMode
		modeUART = 1
		modeNetwork
	End Enum
	
	Public Sub Log_Info(ByRef strLog As String)
		Form1.CheckStep.Text = Form1.CheckStep.Text & strLog & vbCrLf
		Form1.CheckStep.SelectionStart = Len(Form1.CheckStep.Text)
		Form1.CheckStep.Focus()
	End Sub
End Module