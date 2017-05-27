Attribute VB_Name = "Config"
'**********************************************
' Class module for handling config.xml of the
' application.
'**********************************************

Option Explicit
Public gstrXmlPath As String
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Private Type udtConfigData
    strModel As String
    CommMode As CommunicationMode
    strComBaud As String
    intComID As Integer
    strInputSource As String
    lngDelayMs As Long
    intChannelNum As Integer
    intBarCodeLen As Integer
    intLvSpec As Integer
    strVPGModel As String
    strVPGTiming As String
    strVPG100IRE As String
    strVPG80IRE As String
    strVPG20IRE As String
    lngI2cClockRate As Long
    bolEnableCool2 As Boolean
    bolEnableCool1 As Boolean
    bolEnableNormal As Boolean
    bolEnableWarm1 As Boolean
    bolEnableWarm2 As Boolean
    bolEnableChkColor As Boolean
    bolEnableAdjOffset As Boolean
    strChipSet As String
    intSPECCool1x As Integer
    intSPECCool1y As Integer
    intSPECCool1Lv As Integer
    intSPECNormalx As Integer
    intSPECNormaly As Integer
    intSPECNormalLv As Integer
    intSPECWarm1x As Integer
    intSPECWarm1y As Integer
    intSPECWarm1Lv As Integer
    intTOLCool1xt As Integer
    intTOLCool1yt As Integer
    intTOLNormalxt As Integer
    intTOLNormalyt As Integer
    intTOLWarm1xt As Integer
    intTOLWarm1yt As Integer
    intCHKCool1Cxt As Integer
    intCHKCool1Cyt As Integer
    intCHKNormalCxt As Integer
    intCHKNormalCyt As Integer
    intCHKWarm1Cxt As Integer
    intCHKWarm1Cyt As Integer
    intPRESETGANCool1R As Integer
    intPRESETGANCool1G As Integer
    intPRESETGANCool1B As Integer
    intPRESETGANNormalR As Integer
    intPRESETGANNormalG As Integer
    intPRESETGANNormalB As Integer
    intPRESETGANWarm1R As Integer
    intPRESETGANWarm1G As Integer
    intPRESETGANWarm1B As Integer
    intPRESETOFFCool1R As Integer
    intPRESETOFFCool1G As Integer
    intPRESETOFFCool1B As Integer
    intPRESETOFFNormalR As Integer
    intPRESETOFFNormalG As Integer
    intPRESETOFFNormalB As Integer
    intPRESETOFFWarm1R As Integer
    intPRESETOFFWarm1G As Integer
    intPRESETOFFWarm1B As Integer
    intCLEVELRGBGMin As Integer
    intCLEVELRGBGMax As Integer
    intCLEVELRGBOMin As Integer
    intCLEVELRGBOMax As Integer
    intMAGICVALGMin As Integer
    intMAGICVALGMax As Integer
    intMAGICVALOMin As Integer
    intMAGICVALOMax As Integer
End Type

Private mConfigData As udtConfigData

Private Sub Class_Initialize()
    
    
    mConfigData.CommMode = modeUART
    mConfigData.strComBaud = "115200"
    mConfigData.intComID = 1
    mConfigData.strInputSource = "HDMI1"
    mConfigData.lngDelayMs = 500
    mConfigData.intChannelNum = 1
    mConfigData.intBarCodeLen = 1
    mConfigData.intLvSpec = 280
    mConfigData.strVPGModel = "22294"
    mConfigData.strVPGTiming = "69"
    mConfigData.strVPG100IRE = "101"
    mConfigData.strVPG80IRE = "103"
    mConfigData.strVPG20IRE = "109"
    mConfigData.lngI2cClockRate = 50
    mConfigData.bolEnableCool2 = False
    mConfigData.bolEnableCool1 = True
    mConfigData.bolEnableNormal = True
    mConfigData.bolEnableWarm1 = True
    mConfigData.bolEnableWarm2 = False
    mConfigData.bolEnableChkColor = True
    mConfigData.bolEnableAdjOffset = False
    mConfigData.strChipSet = "Null"
End Sub

Public Sub LoadConfigData()

    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(gstrXmlPath)
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        If xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "UART" Then
            mConfigData.CommMode = modeUART
        ElseIf xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "Network" Then
            mConfigData.CommMode = modeNetwork
        ElseIf xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "I2C" Then
            mConfigData.CommMode = modeI2c
        End If
        mConfigData.strModel = xmlDoc.selectSingleNode("/config/model").Text
        gstrCurProjName = mConfigData.strModel
        mConfigData.strComBaud = xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@baud").Text
        mConfigData.intComID = val(xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@id").Text)
        mConfigData.lngI2cClockRate = val(xmlDoc.selectSingleNode("/config/communication/i2c").selectSingleNode("@clockrate").Text)
        mConfigData.strInputSource = xmlDoc.selectSingleNode("/config/input_source").Text
        mConfigData.lngDelayMs = val(xmlDoc.selectSingleNode("/config/delayms").Text)
        mConfigData.intChannelNum = val(xmlDoc.selectSingleNode("/config/channel_number").Text)
        mConfigData.intBarCodeLen = val(xmlDoc.selectSingleNode("/config/length_bar_code").Text)
        mConfigData.intLvSpec = val(xmlDoc.selectSingleNode("/config/Lv_spec").Text)
        mConfigData.strVPGModel = xmlDoc.selectSingleNode("/config/VPG/model").Text
        mConfigData.strVPGTiming = xmlDoc.selectSingleNode("/config/VPG/timing").Text
        mConfigData.strVPG100IRE = xmlDoc.selectSingleNode("/config/VPG/IRE100").Text
        mConfigData.strVPG80IRE = xmlDoc.selectSingleNode("/config/VPG/IRE80").Text
        mConfigData.strVPG20IRE = xmlDoc.selectSingleNode("/config/VPG/IRE20").Text
        mConfigData.strChipSet = xmlDoc.selectSingleNode("/config/chipset").Text
        mConfigData.intSPECCool1x = val(xmlDoc.selectSingleNode("/config/SPEC/cool1/x").Text)
        mConfigData.intSPECCool1y = val(xmlDoc.selectSingleNode("/config/SPEC/cool1/y").Text)
        mConfigData.intSPECCool1Lv = val(xmlDoc.selectSingleNode("/config/SPEC/cool1/Lv").Text)
        mConfigData.intSPECNormalx = val(xmlDoc.selectSingleNode("/config/SPEC/normal/x").Text)
        mConfigData.intSPECNormaly = val(xmlDoc.selectSingleNode("/config/SPEC/normal/y").Text)
        mConfigData.intSPECNormalLv = val(xmlDoc.selectSingleNode("/config/SPEC/normal/Lv").Text)
        mConfigData.intSPECWarm1x = val(xmlDoc.selectSingleNode("/config/SPEC/warm1/x").Text)
        mConfigData.intSPECWarm1y = val(xmlDoc.selectSingleNode("/config/SPEC/warm1/y").Text)
        mConfigData.intSPECWarm1Lv = val(xmlDoc.selectSingleNode("/config/SPEC/warm1/Lv").Text)
        mConfigData.intTOLCool1xt = val(xmlDoc.selectSingleNode("/config/TOL/cool1/xt").Text)
        mConfigData.intTOLCool1yt = val(xmlDoc.selectSingleNode("/config/TOL/cool1/yt").Text)
        mConfigData.intTOLNormalxt = val(xmlDoc.selectSingleNode("/config/TOL/normal/xt").Text)
        mConfigData.intTOLNormalyt = val(xmlDoc.selectSingleNode("/config/TOL/normal/yt").Text)
        mConfigData.intTOLWarm1xt = val(xmlDoc.selectSingleNode("/config/TOL/warm1/xt").Text)
        mConfigData.intTOLWarm1yt = val(xmlDoc.selectSingleNode("/config/TOL/warm1/yt").Text)
        mConfigData.intCHKCool1Cxt = val(xmlDoc.selectSingleNode("/config/CHK/cool1/cxt").Text)
        mConfigData.intCHKCool1Cyt = val(xmlDoc.selectSingleNode("/config/CHK/cool1/cyt").Text)
        mConfigData.intCHKNormalCxt = val(xmlDoc.selectSingleNode("/config/CHK/normal/cxt").Text)
        mConfigData.intCHKNormalCyt = val(xmlDoc.selectSingleNode("/config/CHK/normal/cyt").Text)
        mConfigData.intCHKWarm1Cxt = val(xmlDoc.selectSingleNode("/config/CHK/warm1/cxt").Text)
        mConfigData.intCHKWarm1Cyt = val(xmlDoc.selectSingleNode("/config/CHK/warm1/cyt").Text)
        mConfigData.intPRESETGANCool1R = val(xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/R").Text)
        mConfigData.intPRESETGANCool1G = val(xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/G").Text)
        mConfigData.intPRESETGANCool1B = val(xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/B").Text)
        mConfigData.intPRESETGANNormalR = val(xmlDoc.selectSingleNode("/config/PRESETGAN/normal/R").Text)
        mConfigData.intPRESETGANNormalG = val(xmlDoc.selectSingleNode("/config/PRESETGAN/normal/G").Text)
        mConfigData.intPRESETGANNormalB = val(xmlDoc.selectSingleNode("/config/PRESETGAN/normal/B").Text)
        mConfigData.intPRESETGANWarm1R = val(xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/R").Text)
        mConfigData.intPRESETGANWarm1G = val(xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/G").Text)
        mConfigData.intPRESETGANWarm1B = val(xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/B").Text)
        mConfigData.intPRESETOFFCool1R = val(xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/R").Text)
        mConfigData.intPRESETOFFCool1G = val(xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/G").Text)
        mConfigData.intPRESETOFFCool1B = val(xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/B").Text)
        mConfigData.intPRESETOFFNormalR = val(xmlDoc.selectSingleNode("/config/PRESETOFF/normal/R").Text)
        mConfigData.intPRESETOFFNormalG = val(xmlDoc.selectSingleNode("/config/PRESETOFF/normal/G").Text)
        mConfigData.intPRESETOFFNormalB = val(xmlDoc.selectSingleNode("/config/PRESETOFF/normal/B").Text)
        mConfigData.intPRESETOFFWarm1R = val(xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/R").Text)
        mConfigData.intPRESETOFFWarm1G = val(xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/G").Text)
        mConfigData.intPRESETOFFWarm1B = val(xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/B").Text)
        mConfigData.intCLEVELRGBGMin = val(xmlDoc.selectSingleNode("/config/CLEVELRGB/gain/min").Text)
        mConfigData.intCLEVELRGBGMax = val(xmlDoc.selectSingleNode("/config/CLEVELRGB/gain/max").Text)
        mConfigData.intCLEVELRGBOMin = val(xmlDoc.selectSingleNode("/config/CLEVELRGB/offset/min").Text)
        mConfigData.intCLEVELRGBOMax = val(xmlDoc.selectSingleNode("/config/CLEVELRGB/offset/max").Text)
        mConfigData.intMAGICVALGMin = val(xmlDoc.selectSingleNode("/config/MAGICVAL/x/stepgain").Text)
        mConfigData.intMAGICVALGMax = val(xmlDoc.selectSingleNode("/config/MAGICVAL/x/stepoffset").Text)
        mConfigData.intMAGICVALOMin = val(xmlDoc.selectSingleNode("/config/MAGICVAL/y/stepgain").Text)
        mConfigData.intMAGICVALOMax = val(xmlDoc.selectSingleNode("/config/MAGICVAL/y/stepoffset").Text)
    End If

    If xmlDoc.selectSingleNode("/config/cool_2").Text = "True" Then
        mConfigData.bolEnableCool2 = True
    Else
        mConfigData.bolEnableCool2 = False
    End If
    
    If xmlDoc.selectSingleNode("/config/cool_1").Text = "True" Then
        mConfigData.bolEnableCool1 = True
    Else
        mConfigData.bolEnableCool1 = False
    End If
    
    If xmlDoc.selectSingleNode("/config/normal").Text = "True" Then
        mConfigData.bolEnableNormal = True
    Else
        mConfigData.bolEnableNormal = False
    End If
    
    If xmlDoc.selectSingleNode("/config/warm_1").Text = "True" Then
        mConfigData.bolEnableWarm1 = True
    Else
        mConfigData.bolEnableWarm1 = False
    End If
    
    If xmlDoc.selectSingleNode("/config/warm_2").Text = "True" Then
        mConfigData.bolEnableWarm2 = True
    Else
        mConfigData.bolEnableWarm2 = False
    End If
    
    If xmlDoc.selectSingleNode("/config/check_color").Text = "True" Then
        mConfigData.bolEnableChkColor = True
    Else
        mConfigData.bolEnableChkColor = False
    End If
    
    If xmlDoc.selectSingleNode("/config/adjust_offset").Text = "True" Then
        mConfigData.bolEnableAdjOffset = True
    Else
        mConfigData.bolEnableAdjOffset = False
    End If
End Sub

Public Sub SaveConfigData()
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(gstrXmlPath)
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        If mConfigData.CommMode = modeUART Then
            xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "UART"
        ElseIf mConfigData.CommMode = modeNetwork Then
            xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "Network"
        ElseIf mConfigData.CommMode = modeI2c Then
            xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "I2C"
        End If
        xmlDoc.selectSingleNode("/config/model").Text = mConfigData.strModel
        xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@baud").Text = mConfigData.strComBaud
        xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@id").Text = CStr(mConfigData.intComID)
        xmlDoc.selectSingleNode("/config/input_source").Text = mConfigData.strInputSource
        xmlDoc.selectSingleNode("/config/delayms").Text = CStr(mConfigData.lngDelayMs)
        xmlDoc.selectSingleNode("/config/channel_number").Text = CStr(mConfigData.intChannelNum)
        xmlDoc.selectSingleNode("/config/length_bar_code").Text = CStr(mConfigData.intBarCodeLen)
        xmlDoc.selectSingleNode("/config/Lv_spec").Text = CStr(mConfigData.intLvSpec)
        xmlDoc.selectSingleNode("/config/VPG/model").Text = mConfigData.strVPGModel
        xmlDoc.selectSingleNode("/config/VPG/timing").Text = mConfigData.strVPGTiming
        xmlDoc.selectSingleNode("/config/VPG/IRE100").Text = mConfigData.strVPG100IRE
        xmlDoc.selectSingleNode("/config/VPG/IRE80").Text = mConfigData.strVPG80IRE
        xmlDoc.selectSingleNode("/config/VPG/IRE20").Text = mConfigData.strVPG20IRE
        xmlDoc.selectSingleNode("/config/SPEC/cool1/x").Text = CStr(mConfigData.intSPECCool1x)
        xmlDoc.selectSingleNode("/config/SPEC/cool1/y").Text = CStr(mConfigData.intSPECCool1y)
        xmlDoc.selectSingleNode("/config/SPEC/cool1/Lv").Text = CStr(mConfigData.intSPECCool1Lv)
        xmlDoc.selectSingleNode("/config/SPEC/normal/x").Text = CStr(mConfigData.intSPECNormalx)
        xmlDoc.selectSingleNode("/config/SPEC/normal/y").Text = CStr(mConfigData.intSPECNormaly)
        xmlDoc.selectSingleNode("/config/SPEC/normal/Lv").Text = CStr(mConfigData.intSPECNormalLv)
        xmlDoc.selectSingleNode("/config/SPEC/warm1/x").Text = CStr(mConfigData.intSPECWarm1x)
        xmlDoc.selectSingleNode("/config/SPEC/warm1/y").Text = CStr(mConfigData.intSPECWarm1y)
        xmlDoc.selectSingleNode("/config/SPEC/warm1/Lv").Text = CStr(mConfigData.intSPECWarm1Lv)
        xmlDoc.selectSingleNode("/config/TOL/cool1/xt").Text = CStr(mConfigData.intTOLCool1xt)
        xmlDoc.selectSingleNode("/config/TOL/cool1/yt").Text = CStr(mConfigData.intTOLCool1yt)
        xmlDoc.selectSingleNode("/config/TOL/normal/xt").Text = CStr(mConfigData.intTOLNormalxt)
        xmlDoc.selectSingleNode("/config/TOL/normal/yt").Text = CStr(mConfigData.intTOLNormalyt)
        xmlDoc.selectSingleNode("/config/TOL/warm1/xt").Text = CStr(mConfigData.intTOLWarm1xt)
        xmlDoc.selectSingleNode("/config/TOL/warm1/yt").Text = CStr(mConfigData.intTOLWarm1yt)
        xmlDoc.selectSingleNode("/config/CHK/cool1/cxt").Text = CStr(mConfigData.intCHKCool1Cxt)
        xmlDoc.selectSingleNode("/config/CHK/cool1/cyt").Text = CStr(mConfigData.intCHKCool1Cyt)
        xmlDoc.selectSingleNode("/config/CHK/normal/cxt").Text = CStr(mConfigData.intCHKNormalCxt)
        xmlDoc.selectSingleNode("/config/CHK/normal/cyt").Text = CStr(mConfigData.intCHKNormalCyt)
        xmlDoc.selectSingleNode("/config/CHK/warm1/cxt").Text = CStr(mConfigData.intCHKWarm1Cxt)
        xmlDoc.selectSingleNode("/config/CHK/warm1/cyt").Text = CStr(mConfigData.intCHKWarm1Cyt)
        xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/R").Text = CStr(mConfigData.intPRESETGANCool1R)
        xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/G").Text = CStr(mConfigData.intPRESETGANCool1G)
        xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/B").Text = CStr(mConfigData.intPRESETGANCool1B)
        xmlDoc.selectSingleNode("/config/PRESETGAN/normal/R").Text = CStr(mConfigData.intPRESETGANNormalR)
        xmlDoc.selectSingleNode("/config/PRESETGAN/normal/G").Text = CStr(mConfigData.intPRESETGANNormalG)
        xmlDoc.selectSingleNode("/config/PRESETGAN/normal/B").Text = CStr(mConfigData.intPRESETGANNormalB)
        xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/R").Text = CStr(mConfigData.intPRESETGANWarm1R)
        xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/G").Text = CStr(mConfigData.intPRESETGANWarm1G)
        xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/B").Text = CStr(mConfigData.intPRESETGANWarm1B)
        xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/R").Text = CStr(mConfigData.intPRESETOFFCool1R)
        xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/G").Text = CStr(mConfigData.intPRESETOFFCool1G)
        xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/B").Text = CStr(mConfigData.intPRESETOFFCool1B)
        xmlDoc.selectSingleNode("/config/PRESETOFF/normal/R").Text = CStr(mConfigData.intPRESETOFFNormalR)
        xmlDoc.selectSingleNode("/config/PRESETOFF/normal/G").Text = CStr(mConfigData.intPRESETOFFNormalG)
        xmlDoc.selectSingleNode("/config/PRESETOFFnormal/B").Text = CStr(mConfigData.intPRESETOFFNormalB)
        xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/R").Text = CStr(mConfigData.intPRESETOFFWarm1R)
        xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/G").Text = CStr(mConfigData.intPRESETOFFWarm1G)
        xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/B").Text = CStr(mConfigData.intPRESETOFFWarm1B)
        xmlDoc.selectSingleNode("/config/CLEVELRGB/gain/min").Text = CStr(mConfigData.intCLEVELRGBGMin)
        xmlDoc.selectSingleNode("/config/CLEVELRGB/gain/max").Text = CStr(mConfigData.intCLEVELRGBGMax)
        xmlDoc.selectSingleNode("/config/CLEVELRGB/offset/min").Text = CStr(mConfigData.intCLEVELRGBOMin)
        xmlDoc.selectSingleNode("/config/CLEVELRGB/offset/max").Text = CStr(mConfigData.intCLEVELRGBOMax)
        xmlDoc.selectSingleNode("/config/MAGICVAL/x/stepgain").Text = CStr(mConfigData.intMAGICVALGMin)
        xmlDoc.selectSingleNode("/config/MAGICVAL/x/stepoffset").Text = CStr(mConfigData.intMAGICVALGMax)
        xmlDoc.selectSingleNode("/config/MAGICVAL/y/stepgain").Text = CStr(mConfigData.intMAGICVALOMin)
        xmlDoc.selectSingleNode("/config/MAGICVAL/y/stepoffset").Text = CStr(mConfigData.intMAGICVALOMax)
        
        If mConfigData.bolEnableCool2 Then
            xmlDoc.selectSingleNode("/config/cool_2").Text = "True"
        Else
            xmlDoc.selectSingleNode("/config/cool_2").Text = "False"
        End If
        
        If mConfigData.bolEnableCool1 Then
            xmlDoc.selectSingleNode("/config/cool_1").Text = "True"
        Else
            xmlDoc.selectSingleNode("/config/cool_1").Text = "False"
        End If
        
        If mConfigData.bolEnableNormal Then
            xmlDoc.selectSingleNode("/config/normal").Text = "True"
        Else
            xmlDoc.selectSingleNode("/config/normal").Text = "False"
        End If
        
        If mConfigData.bolEnableWarm1 Then
            xmlDoc.selectSingleNode("/config/warm_1").Text = "True"
        Else
            xmlDoc.selectSingleNode("/config/warm_1").Text = "False"
        End If
        
        If mConfigData.bolEnableWarm2 Then
            xmlDoc.selectSingleNode("/config/warm_2").Text = "True"
        Else
            xmlDoc.selectSingleNode("/config/warm_2").Text = "False"
        End If
        
        If mConfigData.bolEnableChkColor Then
            xmlDoc.selectSingleNode("/config/check_color").Text = "True"
        Else
            xmlDoc.selectSingleNode("/config/check_color").Text = "False"
        End If
        
        If mConfigData.bolEnableAdjOffset Then
            xmlDoc.selectSingleNode("/config/adjust_offset").Text = "True"
        Else
            xmlDoc.selectSingleNode("/config/adjust_offset").Text = "False"
        End If
        
        xmlDoc.Save gstrXmlPath
    End If
End Sub

Public Property Get CommMode() As CommunicationMode
    CommMode = mConfigData.CommMode
End Property

Public Property Let CommMode(enuCommMode As CommunicationMode)
    mConfigData.CommMode = enuCommMode
End Property

Public Property Get ComBaud() As String
    ComBaud = mConfigData.strComBaud
End Property

Public Property Let ComBaud(strComBaud As String)
    mConfigData.strComBaud = strComBaud
End Property

Public Property Get ComID() As Integer
    ComID = mConfigData.intComID
End Property

Public Property Let ComID(intComID As Integer)
    mConfigData.intComID = intComID
End Property

Public Property Get Model() As String
    Model = mConfigData.strModel
End Property

Public Property Let Model(strModel As String)
    mConfigData.strModel = strModel
End Property

Public Property Get inputSource() As String
    inputSource = mConfigData.strInputSource
End Property

Public Property Let inputSource(strInputSource As String)
    mConfigData.strInputSource = strInputSource
End Property

Public Property Get DelayMS() As Long
    DelayMS = mConfigData.lngDelayMs
End Property

Public Property Let DelayMS(lngDelayMs As Long)
    mConfigData.lngDelayMs = lngDelayMs
End Property

Public Property Get ChannelNum() As Integer
    ChannelNum = mConfigData.intChannelNum
End Property

Public Property Let ChannelNum(intChannelNum As Integer)
    mConfigData.intChannelNum = intChannelNum
End Property

Public Property Get BarCodeLen() As Integer
    BarCodeLen = mConfigData.intBarCodeLen
End Property

Public Property Let BarCodeLen(intBarCodeLen As Integer)
    mConfigData.intBarCodeLen = intBarCodeLen
End Property

Public Property Get LvSpec() As Integer
    LvSpec = mConfigData.intLvSpec
End Property

Public Property Let LvSpec(intLvSpec As Integer)
    mConfigData.intLvSpec = intLvSpec
End Property

Public Property Get VPGModel() As String
    VPGModel = mConfigData.strVPGModel
End Property

Public Property Let VPGModel(strVPGModel As String)
    mConfigData.strVPGModel = strVPGModel
End Property

Public Property Get VPGTiming() As String
    VPGTiming = mConfigData.strVPGTiming
End Property

Public Property Let VPGTiming(strVPGTiming As String)
    mConfigData.strVPGTiming = strVPGTiming
End Property

Public Property Get VPG100IRE() As String
    VPG100IRE = mConfigData.strVPG100IRE
End Property

Public Property Let VPG100IRE(strVPG100IRE As String)
    mConfigData.strVPG100IRE = strVPG100IRE
End Property

Public Property Get VPG80IRE() As String
    VPG80IRE = mConfigData.strVPG80IRE
End Property

Public Property Let VPG80IRE(strVPG80IRE As String)
    mConfigData.strVPG80IRE = strVPG80IRE
End Property

Public Property Get VPG20IRE() As String
    VPG20IRE = mConfigData.strVPG20IRE
End Property

Public Property Let VPG20IRE(strVPG20IRE As String)
    mConfigData.strVPG20IRE = strVPG20IRE
End Property

Public Property Get SPECCool1x() As Integer
    SPECCool1x = mConfigData.intSPECCool1x
End Property

Public Property Let SPECCool1x(intSPECCool1x As Integer)
    mConfigData.intSPECCool1x = intSPECCool1x
End Property

Public Property Get SPECCool1y() As Integer
    SPECCool1y = mConfigData.intSPECCool1y
End Property

Public Property Let SPECCool1y(intSPECCool1y As Integer)
    mConfigData.intSPECCool1y = intSPECCool1y
End Property

Public Property Get SPECCool1Lv() As Integer
    SPECCool1Lv = mConfigData.intSPECCool1Lv
End Property

Public Property Let SPECCool1Lv(intSPECCool1Lv As Integer)
    mConfigData.intSPECCool1Lv = intSPECCool1Lv
End Property

Public Property Get SPECNormalx() As Integer
    SPECNormalx = mConfigData.intSPECNormalx
End Property

Public Property Let SPECNormalx(intSPECNormalx As Integer)
    mConfigData.intSPECNormalx = intSPECNormalx
End Property

Public Property Get SPECNormaly() As Integer
    SPECNormaly = mConfigData.intSPECNormaly
End Property

Public Property Let SPECNormaly(intSPECNormaly As Integer)
    mConfigData.intSPECNormaly = intSPECNormaly
End Property

Public Property Get SPECNormalLv() As Integer
    SPECNormalLv = mConfigData.intSPECNormalLv
End Property

Public Property Let SPECNormalLv(intSPECNormalLv As Integer)
    mConfigData.intSPECNormalLv = intSPECNormalLv
End Property

Public Property Get SPECWarm1x() As Integer
    SPECWarm1x = mConfigData.intSPECWarm1x
End Property

Public Property Let SPECWarm1x(intSPECWarm1x As Integer)
    mConfigData.intSPECWarm1x = intSPECWarm1x
End Property

Public Property Get SPECWarm1y() As Integer
    SPECWarm1y = mConfigData.intSPECWarm1y
End Property

Public Property Let SPECWarm1y(intSPECWarm1y As Integer)
    mConfigData.intSPECWarm1y = intSPECWarm1y
End Property

Public Property Get SPECWarm1Lv() As Integer
    SPECWarm1Lv = mConfigData.intSPECWarm1Lv
End Property

Public Property Let SPECWarm1Lv(intSPECWarm1Lv As Integer)
    mConfigData.intSPECWarm1Lv = intSPECWarm1Lv
End Property

Public Property Get TOLCool1xt() As Integer
    TOLCool1xt = mConfigData.intTOLCool1xt
End Property

Public Property Let TOLCool1xt(intTOLCool1xt As Integer)
    mConfigData.intTOLCool1xt = intTOLCool1xt
End Property

Public Property Get TOLCool1yt() As Integer
    TOLCool1yt = mConfigData.intTOLCool1yt
End Property

Public Property Let TOLCool1yt(intTOLCool1yt As Integer)
    mConfigData.intTOLCool1yt = intTOLCool1yt
End Property

Public Property Get TOLNormalxt() As Integer
    TOLNormalxt = mConfigData.intTOLNormalxt
End Property

Public Property Let TOLNormalxt(intTOLNormalxt As Integer)
    mConfigData.intTOLNormalxt = intTOLNormalxt
End Property

Public Property Get TOLNormalyt() As Integer
    TOLNormalyt = mConfigData.intTOLNormalyt
End Property

Public Property Let TOLNormalyt(intTOLNormalyt As Integer)
    mConfigData.intTOLNormalyt = intTOLNormalyt
End Property

Public Property Get TOLWarm1xt() As Integer
    TOLWarm1xt = mConfigData.intTOLWarm1xt
End Property

Public Property Let TOLWarm1xt(intTOLWarm1xt As Integer)
    mConfigData.intTOLWarm1xt = intTOLWarm1xt
End Property

Public Property Get TOLWarm1yt() As Integer
    TOLWarm1yt = mConfigData.intTOLWarm1yt
End Property

Public Property Let TOLWarm1yt(intTOLWarm1yt As Integer)
    mConfigData.intTOLWarm1yt = intTOLWarm1yt
End Property

Public Property Get CHKCool1Cxt() As Integer
    CHKCool1Cxt = mConfigData.intCHKCool1Cxt
End Property

Public Property Let CHKCool1Cxt(intCHKCool1Cxt As Integer)
    mConfigData.intCHKCool1Cxt = intCHKCool1Cxt
End Property

Public Property Get CHKCool1Cyt() As Integer
    CHKCool1Cyt = mConfigData.intCHKCool1Cyt
End Property

Public Property Let CHKCool1Cyt(intCHKCool1Cyt As Integer)
    mConfigData.intCHKCool1Cyt = intCHKCool1Cyt
End Property

Public Property Get CHKNormalCxt() As Integer
    CHKNormalCxt = mConfigData.intCHKNormalCxt
End Property

Public Property Let CHKNormalCxt(intCHKNormalCxt As Integer)
    mConfigData.intCHKNormalCxt = intCHKNormalCxt
End Property

Public Property Get CHKNormalCyt() As Integer
    CHKNormalCyt = mConfigData.intCHKNormalCyt
End Property

Public Property Let CHKNormalCyt(intCHKNormalCyt As Integer)
    mConfigData.intCHKNormalCyt = intCHKNormalCyt
End Property

Public Property Get CHKWarm1Cxt() As Integer
    CHKWarm1Cxt = mConfigData.intCHKWarm1Cxt
End Property

Public Property Let CHKWarm1Cxt(intCHKWarm1Cxt As Integer)
    mConfigData.intCHKWarm1Cxt = intCHKWarm1Cxt
End Property

Public Property Get CHKWarm1Cyt() As Integer
    CHKWarm1Cyt = mConfigData.intCHKWarm1Cyt
End Property

Public Property Let CHKWarm1Cyt(intCHKWarm1Cyt As Integer)
    mConfigData.intCHKWarm1Cyt = intCHKWarm1Cyt
End Property

Public Property Get PRESETGANCool1R() As Integer
    PRESETGANCool1R = mConfigData.intPRESETGANCool1R
End Property

Public Property Let PRESETGANCool1R(intPRESETGANCool1R As Integer)
    mConfigData.intPRESETGANCool1R = intPRESETGANCool1R
End Property

Public Property Get PRESETGANCool1G() As Integer
    PRESETGANCool1G = mConfigData.intPRESETGANCool1G
End Property

Public Property Let PRESETGANCool1G(intPRESETGANCool1G As Integer)
    mConfigData.intPRESETGANCool1G = intPRESETGANCool1G
End Property

Public Property Get PRESETGANCool1B() As Integer
    PRESETGANCool1B = mConfigData.intPRESETGANCool1B
End Property

Public Property Let PRESETGANCool1B(intPRESETGANCool1B As Integer)
    mConfigData.intPRESETGANCool1B = intPRESETGANCool1B
End Property

Public Property Get PRESETGANNormalR() As Integer
    PRESETGANNormalR = mConfigData.intPRESETGANNormalR
End Property

Public Property Let PRESETGANNormalR(intPRESETGANNormalR As Integer)
    mConfigData.intPRESETGANNormalR = intPRESETGANNormalR
End Property

Public Property Get PRESETGANNormalG() As Integer
    PRESETGANNormalG = mConfigData.intPRESETGANNormalG
End Property

Public Property Let PRESETGANNormalG(intPRESETGANNormalG As Integer)
    mConfigData.intPRESETGANNormalG = intPRESETGANNormalG
End Property

Public Property Get PRESETGANNormalB() As Integer
    PRESETGANNormalB = mConfigData.intPRESETGANNormalB
End Property

Public Property Let PRESETGANNormalB(intPRESETGANNormalB As Integer)
    mConfigData.intPRESETGANNormalB = intPRESETGANNormalB
End Property

Public Property Get PRESETGANWarm1R() As Integer
    PRESETGANWarm1R = mConfigData.intPRESETGANWarm1R
End Property

Public Property Let PRESETGANWarm1R(intPRESETGANWarm1R As Integer)
    mConfigData.intPRESETGANWarm1R = intPRESETGANWarm1R
End Property

Public Property Get PRESETGANWarm1G() As Integer
    PRESETGANWarm1G = mConfigData.intPRESETGANWarm1G
End Property

Public Property Let PRESETGANWarm1G(intPRESETGANWarm1G As Integer)
    mConfigData.intPRESETGANWarm1G = intPRESETGANWarm1G
End Property

Public Property Get PRESETGANWarm1B() As Integer
    PRESETGANWarm1B = mConfigData.intPRESETGANWarm1B
End Property

Public Property Let PRESETGANWarm1B(intPRESETGANWarm1B As Integer)
    mConfigData.intPRESETGANWarm1B = intPRESETGANWarm1B
End Property

Public Property Get PRESETOFFCool1R() As Integer
    PRESETOFFCool1R = mConfigData.intPRESETOFFCool1R
End Property

Public Property Let PRESETOFFCool1R(intPRESETOFFCool1R As Integer)
    mConfigData.intPRESETOFFCool1R = intPRESETOFFCool1R
End Property

Public Property Get PRESETOFFCool1G() As Integer
    PRESETOFFCool1G = mConfigData.intPRESETOFFCool1G
End Property

Public Property Let PRESETOFFCool1G(intPRESETOFFCool1G As Integer)
    mConfigData.intPRESETOFFCool1G = intPRESETOFFCool1G
End Property

Public Property Get PRESETOFFCool1B() As Integer
    PRESETOFFCool1B = mConfigData.intPRESETOFFCool1B
End Property

Public Property Let PRESETOFFCool1B(intPRESETOFFCool1B As Integer)
    mConfigData.intPRESETOFFCool1B = intPRESETOFFCool1B
End Property

Public Property Get PRESETOFFNormalR() As Integer
    PRESETOFFNormalR = mConfigData.intPRESETOFFNormalR
End Property

Public Property Let PRESETOFFNormalR(intPRESETOFFNormalR As Integer)
    mConfigData.intPRESETOFFNormalR = intPRESETOFFNormalR
End Property

Public Property Get PRESETOFFNormalG() As Integer
    PRESETOFFNormalG = mConfigData.intPRESETOFFNormalG
End Property

Public Property Let PRESETOFFNormalG(intPRESETOFFNormalG As Integer)
    mConfigData.intPRESETOFFNormalG = intPRESETOFFNormalG
End Property

Public Property Get PRESETOFFNormalB() As Integer
    PRESETOFFNormalB = mConfigData.intPRESETOFFNormalB
End Property

Public Property Let PRESETOFFNormalB(intPRESETOFFNormalB As Integer)
    mConfigData.intPRESETOFFNormalB = intPRESETOFFNormalB
End Property

Public Property Get PRESETOFFWarm1R() As Integer
    PRESETOFFWarm1R = mConfigData.intPRESETOFFWarm1R
End Property

Public Property Let PRESETOFFWarm1R(intPRESETOFFWarm1R As Integer)
    mConfigData.intPRESETOFFWarm1R = intPRESETOFFWarm1R
End Property

Public Property Get PRESETOFFWarm1G() As Integer
    PRESETOFFWarm1G = mConfigData.intPRESETOFFWarm1G
End Property

Public Property Let PRESETOFFWarm1G(intPRESETOFFWarm1G As Integer)
    mConfigData.intPRESETOFFWarm1G = intPRESETOFFWarm1G
End Property

Public Property Get PRESETOFFWarm1B() As Integer
    PRESETOFFWarm1B = mConfigData.intPRESETOFFWarm1B
End Property

Public Property Let PRESETOFFWarm1B(intPRESETOFFWarm1B As Integer)
    mConfigData.intPRESETOFFWarm1B = intPRESETOFFWarm1B
End Property

Public Property Get CLEVELRGBGMin() As Integer
    CLEVELRGBGMin = mConfigData.intCLEVELRGBGMin
End Property

Public Property Let CLEVELRGBGMin(intCLEVELRGBGMin As Integer)
    mConfigData.intCLEVELRGBGMin = intCLEVELRGBGMin
End Property

Public Property Get CLEVELRGBGMax() As Integer
    CLEVELRGBGMax = mConfigData.intCLEVELRGBGMax
End Property

Public Property Let CLEVELRGBGMax(intCLEVELRGBGMax As Integer)
    mConfigData.intCLEVELRGBGMax = intCLEVELRGBGMax
End Property

Public Property Get CLEVELRGBOMin() As Integer
    CLEVELRGBOMin = mConfigData.intCLEVELRGBOMin
End Property

Public Property Let CLEVELRGBOMin(intCLEVELRGBOMin As Integer)
    mConfigData.intCLEVELRGBOMin = intCLEVELRGBOMin
End Property

Public Property Get CLEVELRGBOMax() As Integer
    CLEVELRGBOMax = mConfigData.intCLEVELRGBOMax
End Property

Public Property Let CLEVELRGBOMax(intCLEVELRGBOMax As Integer)
    mConfigData.intCLEVELRGBOMax = intCLEVELRGBOMax
End Property

Public Property Get MAGICVALGMin() As Integer
    MAGICVALGMin = mConfigData.intMAGICVALGMin
End Property

Public Property Let MAGICVALGMin(intMAGICVALGMin As Integer)
    mConfigData.intMAGICVALGMin = intMAGICVALGMin
End Property

Public Property Get MAGICVALGMax() As Integer
    MAGICVALGMax = mConfigData.intMAGICVALGMax
End Property

Public Property Let MAGICVALGMax(intMAGICVALGMax As Integer)
    mConfigData.intMAGICVALGMax = intMAGICVALGMax
End Property

Public Property Get MAGICVALOMin() As Integer
    MAGICVALOMin = mConfigData.intMAGICVALOMin
End Property

Public Property Let MAGICVALOMin(intMAGICVALOMin As Integer)
    mConfigData.intMAGICVALOMin = intMAGICVALOMin
End Property

Public Property Get MAGICVALOMax() As Integer
    MAGICVALOMax = mConfigData.intMAGICVALOMax
End Property

Public Property Let MAGICVALOMax(intMAGICVALOMax As Integer)
    mConfigData.intMAGICVALOMax = intMAGICVALOMax
End Property


Public Property Get I2cClockRate() As String
    I2cClockRate = mConfigData.lngI2cClockRate
End Property

Public Property Let I2cClockRate(lngI2cClockRate As String)
    mConfigData.lngI2cClockRate = lngI2cClockRate
End Property

Public Property Get EnableCool2() As Boolean
    EnableCool2 = mConfigData.bolEnableCool2
End Property

Public Property Let EnableCool2(bolEnableCool2 As Boolean)
    mConfigData.bolEnableCool2 = bolEnableCool2
End Property

Public Property Get EnableCool1() As Boolean
    EnableCool1 = mConfigData.bolEnableCool1
End Property

Public Property Let EnableCool1(bolEnableCool1 As Boolean)
    mConfigData.bolEnableCool1 = bolEnableCool1
End Property

Public Property Get EnableNormal() As Boolean
    EnableNormal = mConfigData.bolEnableNormal
End Property

Public Property Let EnableNormal(bolEnableNormal As Boolean)
    mConfigData.bolEnableNormal = bolEnableNormal
End Property

Public Property Get EnableWarm1() As Boolean
    EnableWarm1 = mConfigData.bolEnableWarm1
End Property

Public Property Let EnableWarm1(bolEnableWarm1 As Boolean)
    mConfigData.bolEnableWarm1 = bolEnableWarm1
End Property

Public Property Get EnableWarm2() As Boolean
    EnableWarm2 = mConfigData.bolEnableWarm2
End Property

Public Property Let EnableWarm2(bolEnableWarm2 As Boolean)
    mConfigData.bolEnableWarm2 = bolEnableWarm2
End Property

Public Property Get EnableChkColor() As Boolean
    EnableChkColor = mConfigData.bolEnableChkColor
End Property

Public Property Let EnableChkColor(bolEnableChkColor As Boolean)
    mConfigData.bolEnableChkColor = bolEnableChkColor
End Property

Public Property Get EnableAdjOffset() As Boolean
    EnableAdjOffset = mConfigData.bolEnableAdjOffset
End Property

Public Property Let EnableAdjOffset(bolEnableAdjOffset As Boolean)
    mConfigData.bolEnableAdjOffset = bolEnableAdjOffset
End Property

Public Property Get ChipSet() As String
    ChipSet = mConfigData.strChipSet
End Property

