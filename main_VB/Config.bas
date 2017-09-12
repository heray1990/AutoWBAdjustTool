Attribute VB_Name = "Config"
'**********************************************
' Class module for handling config.xml of the
' application.
'**********************************************

Option Explicit
Public gstrXmlPath As String
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long


Private Sub Class_Initialize()
    gudtConfigData.CommMode = modeUART
    gudtConfigData.strComBaud = "115200"
    gudtConfigData.intComID = 1
    gudtConfigData.strInputSource = "HDMI1"
    gudtConfigData.lngDelayMs = 500
    gudtConfigData.intChannelNum = 1
    gudtConfigData.intBarCodeLen = 1
    gudtConfigData.intLvSpec = 280
    gudtConfigData.strVPGModel = "22294"
    gudtConfigData.strVPGTiming = "69"
    gudtConfigData.strVPG100IRE = "101"
    gudtConfigData.strVPG80IRE = "103"
    gudtConfigData.strVPG20IRE = "109"
    gudtConfigData.lngI2cClockRate = 50
    gudtConfigData.bolEnableCool2 = False
    gudtConfigData.bolEnableCool1 = True
    gudtConfigData.bolEnableNormal = True
    gudtConfigData.bolEnableWarm1 = True
    gudtConfigData.bolEnableWarm2 = False
    gudtConfigData.bolEnableChkColor = True
    gudtConfigData.bolEnableAdjOffset = False
    gudtConfigData.strChipSet = "Null"
End Sub

Public Sub LoadConfigData()
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(gstrXmlPath)
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        If xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "UART" Then
            gudtConfigData.CommMode = modeUART
        ElseIf xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "Network" Then
            gudtConfigData.CommMode = modeNetwork
        ElseIf xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "I2C" Then
            gudtConfigData.CommMode = modeI2c
        End If
        gudtConfigData.strModel = xmlDoc.selectSingleNode("/config/model").Text
        gstrCurProjName = gudtConfigData.strModel
        gudtConfigData.strComBaud = xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@baud").Text
        gudtConfigData.intComID = val(xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@id").Text)
        gudtConfigData.lngI2cClockRate = val(xmlDoc.selectSingleNode("/config/communication/i2c").selectSingleNode("@clockrate").Text)
        gudtConfigData.strInputSource = xmlDoc.selectSingleNode("/config/input_source").Text
        gudtConfigData.lngDelayMs = val(xmlDoc.selectSingleNode("/config/delayms").Text)
        gudtConfigData.intChannelNum = val(xmlDoc.selectSingleNode("/config/channel_number").Text)
        gudtConfigData.intBarCodeLen = val(xmlDoc.selectSingleNode("/config/length_bar_code").Text)
        gudtConfigData.intLvSpec = val(xmlDoc.selectSingleNode("/config/Lv_spec").Text)
        gudtConfigData.strVPGModel = xmlDoc.selectSingleNode("/config/VPG/model").Text
        gudtConfigData.strVPGTiming = xmlDoc.selectSingleNode("/config/VPG/timing").Text
        gudtConfigData.strVPG100IRE = xmlDoc.selectSingleNode("/config/VPG/IRE100").Text
        gudtConfigData.strVPG80IRE = xmlDoc.selectSingleNode("/config/VPG/IRE80").Text
        gudtConfigData.strVPG20IRE = xmlDoc.selectSingleNode("/config/VPG/IRE20").Text
        gudtConfigData.strChipSet = xmlDoc.selectSingleNode("/config/chipset").Text
    End If

    If xmlDoc.selectSingleNode("/config/cool_2").Text = "True" Then
        gudtConfigData.bolEnableCool2 = True
    Else
        gudtConfigData.bolEnableCool2 = False
    End If
    
    If xmlDoc.selectSingleNode("/config/cool_1").Text = "True" Then
        gudtConfigData.bolEnableCool1 = True
    Else
        gudtConfigData.bolEnableCool1 = False
    End If
    
    If xmlDoc.selectSingleNode("/config/normal").Text = "True" Then
        gudtConfigData.bolEnableNormal = True
    Else
        gudtConfigData.bolEnableNormal = False
    End If
    
    If xmlDoc.selectSingleNode("/config/warm_1").Text = "True" Then
        gudtConfigData.bolEnableWarm1 = True
    Else
        gudtConfigData.bolEnableWarm1 = False
    End If
    
    If xmlDoc.selectSingleNode("/config/warm_2").Text = "True" Then
        gudtConfigData.bolEnableWarm2 = True
    Else
        gudtConfigData.bolEnableWarm2 = False
    End If
    
    If xmlDoc.selectSingleNode("/config/check_color").Text = "True" Then
        gudtConfigData.bolEnableChkColor = True
    Else
        gudtConfigData.bolEnableChkColor = False
    End If
    
    If xmlDoc.selectSingleNode("/config/adjust_offset").Text = "True" Then
        gudtConfigData.bolEnableAdjOffset = True
    Else
        gudtConfigData.bolEnableAdjOffset = False
    End If
End Sub

Public Sub LoadSpecData()
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(gstrXmlPath)
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        gudtSpecData.intSPECCool1x = val(xmlDoc.selectSingleNode("/config/SPEC/cool1/x").Text)
        gudtSpecData.intSPECCool1y = val(xmlDoc.selectSingleNode("/config/SPEC/cool1/y").Text)
        gudtSpecData.intSPECCool1Lv = val(xmlDoc.selectSingleNode("/config/SPEC/cool1/Lv").Text)
        gudtSpecData.intSPECNormalx = val(xmlDoc.selectSingleNode("/config/SPEC/normal/x").Text)
        gudtSpecData.intSPECNormaly = val(xmlDoc.selectSingleNode("/config/SPEC/normal/y").Text)
        gudtSpecData.intSPECNormalLv = val(xmlDoc.selectSingleNode("/config/SPEC/normal/Lv").Text)
        gudtSpecData.intSPECWarm1x = val(xmlDoc.selectSingleNode("/config/SPEC/warm1/x").Text)
        gudtSpecData.intSPECWarm1y = val(xmlDoc.selectSingleNode("/config/SPEC/warm1/y").Text)
        gudtSpecData.intSPECWarm1Lv = val(xmlDoc.selectSingleNode("/config/SPEC/warm1/Lv").Text)
        gudtSpecData.intTOLCool1xt = val(xmlDoc.selectSingleNode("/config/TOL/cool1/xt").Text)
        gudtSpecData.intTOLCool1yt = val(xmlDoc.selectSingleNode("/config/TOL/cool1/yt").Text)
        gudtSpecData.intTOLNormalxt = val(xmlDoc.selectSingleNode("/config/TOL/normal/xt").Text)
        gudtSpecData.intTOLNormalyt = val(xmlDoc.selectSingleNode("/config/TOL/normal/yt").Text)
        gudtSpecData.intTOLWarm1xt = val(xmlDoc.selectSingleNode("/config/TOL/warm1/xt").Text)
        gudtSpecData.intTOLWarm1yt = val(xmlDoc.selectSingleNode("/config/TOL/warm1/yt").Text)
        gudtSpecData.intCHKCool1Cxt = val(xmlDoc.selectSingleNode("/config/CHK/cool1/cxt").Text)
        gudtSpecData.intCHKCool1Cyt = val(xmlDoc.selectSingleNode("/config/CHK/cool1/cyt").Text)
        gudtSpecData.intCHKNormalCxt = val(xmlDoc.selectSingleNode("/config/CHK/normal/cxt").Text)
        gudtSpecData.intCHKNormalCyt = val(xmlDoc.selectSingleNode("/config/CHK/normal/cyt").Text)
        gudtSpecData.intCHKWarm1Cxt = val(xmlDoc.selectSingleNode("/config/CHK/warm1/cxt").Text)
        gudtSpecData.intCHKWarm1Cyt = val(xmlDoc.selectSingleNode("/config/CHK/warm1/cyt").Text)
        gudtSpecData.intPRESETGANCool1R = val(xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/R").Text)
        gudtSpecData.intPRESETGANCool1G = val(xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/G").Text)
        gudtSpecData.intPRESETGANCool1B = val(xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/B").Text)
        gudtSpecData.intPRESETGANNormalR = val(xmlDoc.selectSingleNode("/config/PRESETGAN/normal/R").Text)
        gudtSpecData.intPRESETGANNormalG = val(xmlDoc.selectSingleNode("/config/PRESETGAN/normal/G").Text)
        gudtSpecData.intPRESETGANNormalB = val(xmlDoc.selectSingleNode("/config/PRESETGAN/normal/B").Text)
        gudtSpecData.intPRESETGANWarm1R = val(xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/R").Text)
        gudtSpecData.intPRESETGANWarm1G = val(xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/G").Text)
        gudtSpecData.intPRESETGANWarm1B = val(xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/B").Text)
        gudtSpecData.intPRESETOFFCool1R = val(xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/R").Text)
        gudtSpecData.intPRESETOFFCool1G = val(xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/G").Text)
        gudtSpecData.intPRESETOFFCool1B = val(xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/B").Text)
        gudtSpecData.intPRESETOFFNormalR = val(xmlDoc.selectSingleNode("/config/PRESETOFF/normal/R").Text)
        gudtSpecData.intPRESETOFFNormalG = val(xmlDoc.selectSingleNode("/config/PRESETOFF/normal/G").Text)
        gudtSpecData.intPRESETOFFNormalB = val(xmlDoc.selectSingleNode("/config/PRESETOFF/normal/B").Text)
        gudtSpecData.intPRESETOFFWarm1R = val(xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/R").Text)
        gudtSpecData.intPRESETOFFWarm1G = val(xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/G").Text)
        gudtSpecData.intPRESETOFFWarm1B = val(xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/B").Text)
        gudtSpecData.intCLEVELRGBGMin = val(xmlDoc.selectSingleNode("/config/CLEVELRGB/gain/min").Text)
        gudtSpecData.intCLEVELRGBGMax = val(xmlDoc.selectSingleNode("/config/CLEVELRGB/gain/max").Text)
        gudtSpecData.intCLEVELRGBOMin = val(xmlDoc.selectSingleNode("/config/CLEVELRGB/offset/min").Text)
        gudtSpecData.intCLEVELRGBOMax = val(xmlDoc.selectSingleNode("/config/CLEVELRGB/offset/max").Text)
        gudtSpecData.intMAGICVALGMin = val(xmlDoc.selectSingleNode("/config/MAGICVAL/x/stepgain").Text)
        gudtSpecData.intMAGICVALGMax = val(xmlDoc.selectSingleNode("/config/MAGICVAL/x/stepoffset").Text)
        gudtSpecData.intMAGICVALOMin = val(xmlDoc.selectSingleNode("/config/MAGICVAL/y/stepgain").Text)
        gudtSpecData.intMAGICVALOMax = val(xmlDoc.selectSingleNode("/config/MAGICVAL/y/stepoffset").Text)
    End If
End Sub

Public Sub SaveConfigData()
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(gstrXmlPath)
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        If gudtConfigData.CommMode = modeUART Then
            xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "UART"
        ElseIf gudtConfigData.CommMode = modeNetwork Then
            xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "Network"
        ElseIf gudtConfigData.CommMode = modeI2c Then
            xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "I2C"
        End If
        xmlDoc.selectSingleNode("/config/model").Text = gudtConfigData.strModel
        xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@baud").Text = gudtConfigData.strComBaud
        xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@id").Text = CStr(gudtConfigData.intComID)
        xmlDoc.selectSingleNode("/config/input_source").Text = gudtConfigData.strInputSource
        xmlDoc.selectSingleNode("/config/delayms").Text = CStr(gudtConfigData.lngDelayMs)
        xmlDoc.selectSingleNode("/config/channel_number").Text = CStr(gudtConfigData.intChannelNum)
        xmlDoc.selectSingleNode("/config/length_bar_code").Text = CStr(gudtConfigData.intBarCodeLen)
        xmlDoc.selectSingleNode("/config/Lv_spec").Text = CStr(gudtConfigData.intLvSpec)
        xmlDoc.selectSingleNode("/config/VPG/model").Text = gudtConfigData.strVPGModel
        xmlDoc.selectSingleNode("/config/VPG/timing").Text = gudtConfigData.strVPGTiming
        xmlDoc.selectSingleNode("/config/VPG/IRE100").Text = gudtConfigData.strVPG100IRE
        xmlDoc.selectSingleNode("/config/VPG/IRE80").Text = gudtConfigData.strVPG80IRE
        xmlDoc.selectSingleNode("/config/VPG/IRE20").Text = gudtConfigData.strVPG20IRE
        
        If gudtConfigData.bolEnableCool2 Then
            xmlDoc.selectSingleNode("/config/cool_2").Text = "True"
        Else
            xmlDoc.selectSingleNode("/config/cool_2").Text = "False"
        End If
        
        If gudtConfigData.bolEnableCool1 Then
            xmlDoc.selectSingleNode("/config/cool_1").Text = "True"
        Else
            xmlDoc.selectSingleNode("/config/cool_1").Text = "False"
        End If
        
        If gudtConfigData.bolEnableNormal Then
            xmlDoc.selectSingleNode("/config/normal").Text = "True"
        Else
            xmlDoc.selectSingleNode("/config/normal").Text = "False"
        End If
        
        If gudtConfigData.bolEnableWarm1 Then
            xmlDoc.selectSingleNode("/config/warm_1").Text = "True"
        Else
            xmlDoc.selectSingleNode("/config/warm_1").Text = "False"
        End If
        
        If gudtConfigData.bolEnableWarm2 Then
            xmlDoc.selectSingleNode("/config/warm_2").Text = "True"
        Else
            xmlDoc.selectSingleNode("/config/warm_2").Text = "False"
        End If
        
        If gudtConfigData.bolEnableChkColor Then
            xmlDoc.selectSingleNode("/config/check_color").Text = "True"
        Else
            xmlDoc.selectSingleNode("/config/check_color").Text = "False"
        End If
        
        If gudtConfigData.bolEnableAdjOffset Then
            xmlDoc.selectSingleNode("/config/adjust_offset").Text = "True"
        Else
            xmlDoc.selectSingleNode("/config/adjust_offset").Text = "False"
        End If
        
        xmlDoc.save gstrXmlPath
    End If
End Sub

Public Sub SaveConfigData1()
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(gstrXmlPath)
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        xmlDoc.selectSingleNode("/config/SPEC/cool1/x").Text = CStr(gudtSpecData.intSPECCool1x)
        xmlDoc.selectSingleNode("/config/SPEC/cool1/y").Text = CStr(gudtSpecData.intSPECCool1y)
        xmlDoc.selectSingleNode("/config/SPEC/cool1/Lv").Text = CStr(gudtSpecData.intSPECCool1Lv)
        xmlDoc.selectSingleNode("/config/SPEC/normal/x").Text = CStr(gudtSpecData.intSPECNormalx)
        xmlDoc.selectSingleNode("/config/SPEC/normal/y").Text = CStr(gudtSpecData.intSPECNormaly)
        xmlDoc.selectSingleNode("/config/SPEC/normal/Lv").Text = CStr(gudtSpecData.intSPECNormalLv)
        xmlDoc.selectSingleNode("/config/SPEC/warm1/x").Text = CStr(gudtSpecData.intSPECWarm1x)
        xmlDoc.selectSingleNode("/config/SPEC/warm1/y").Text = CStr(gudtSpecData.intSPECWarm1y)
        xmlDoc.selectSingleNode("/config/SPEC/warm1/Lv").Text = CStr(gudtSpecData.intSPECWarm1Lv)
        xmlDoc.selectSingleNode("/config/TOL/cool1/xt").Text = CStr(gudtSpecData.intTOLCool1xt)
        xmlDoc.selectSingleNode("/config/TOL/cool1/yt").Text = CStr(gudtSpecData.intTOLCool1yt)
        xmlDoc.selectSingleNode("/config/TOL/normal/xt").Text = CStr(gudtSpecData.intTOLNormalxt)
        xmlDoc.selectSingleNode("/config/TOL/normal/yt").Text = CStr(gudtSpecData.intTOLNormalyt)
        xmlDoc.selectSingleNode("/config/TOL/warm1/xt").Text = CStr(gudtSpecData.intTOLWarm1xt)
        xmlDoc.selectSingleNode("/config/TOL/warm1/yt").Text = CStr(gudtSpecData.intTOLWarm1yt)
        xmlDoc.selectSingleNode("/config/CHK/cool1/cxt").Text = CStr(gudtSpecData.intCHKCool1Cxt)
        xmlDoc.selectSingleNode("/config/CHK/cool1/cyt").Text = CStr(gudtSpecData.intCHKCool1Cyt)
        xmlDoc.selectSingleNode("/config/CHK/normal/cxt").Text = CStr(gudtSpecData.intCHKNormalCxt)
        xmlDoc.selectSingleNode("/config/CHK/normal/cyt").Text = CStr(gudtSpecData.intCHKNormalCyt)
        xmlDoc.selectSingleNode("/config/CHK/warm1/cxt").Text = CStr(gudtSpecData.intCHKWarm1Cxt)
        xmlDoc.selectSingleNode("/config/CHK/warm1/cyt").Text = CStr(gudtSpecData.intCHKWarm1Cyt)
        xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/R").Text = CStr(gudtSpecData.intPRESETGANCool1R)
        xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/G").Text = CStr(gudtSpecData.intPRESETGANCool1G)
        xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/B").Text = CStr(gudtSpecData.intPRESETGANCool1B)
        xmlDoc.selectSingleNode("/config/PRESETGAN/normal/R").Text = CStr(gudtSpecData.intPRESETGANNormalR)
        xmlDoc.selectSingleNode("/config/PRESETGAN/normal/G").Text = CStr(gudtSpecData.intPRESETGANNormalG)
        xmlDoc.selectSingleNode("/config/PRESETGAN/normal/B").Text = CStr(gudtSpecData.intPRESETGANNormalB)
        xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/R").Text = CStr(gudtSpecData.intPRESETGANWarm1R)
        xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/G").Text = CStr(gudtSpecData.intPRESETGANWarm1G)
        xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/B").Text = CStr(gudtSpecData.intPRESETGANWarm1B)
        xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/R").Text = CStr(gudtSpecData.intPRESETOFFCool1R)
        xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/G").Text = CStr(gudtSpecData.intPRESETOFFCool1G)
        xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/B").Text = CStr(gudtSpecData.intPRESETOFFCool1B)
        xmlDoc.selectSingleNode("/config/PRESETOFF/normal/R").Text = CStr(gudtSpecData.intPRESETOFFNormalR)
        xmlDoc.selectSingleNode("/config/PRESETOFF/normal/G").Text = CStr(gudtSpecData.intPRESETOFFNormalG)
        xmlDoc.selectSingleNode("/config/PRESETOFF/normal/B").Text = CStr(gudtSpecData.intPRESETOFFNormalB)
        xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/R").Text = CStr(gudtSpecData.intPRESETOFFWarm1R)
        xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/G").Text = CStr(gudtSpecData.intPRESETOFFWarm1G)
        xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/B").Text = CStr(gudtSpecData.intPRESETOFFWarm1B)
        xmlDoc.selectSingleNode("/config/CLEVELRGB/gain/min").Text = CStr(gudtSpecData.intCLEVELRGBGMin)
        xmlDoc.selectSingleNode("/config/CLEVELRGB/gain/max").Text = CStr(gudtSpecData.intCLEVELRGBGMax)
        xmlDoc.selectSingleNode("/config/CLEVELRGB/offset/min").Text = CStr(gudtSpecData.intCLEVELRGBOMin)
        xmlDoc.selectSingleNode("/config/CLEVELRGB/offset/max").Text = CStr(gudtSpecData.intCLEVELRGBOMax)
        xmlDoc.selectSingleNode("/config/MAGICVAL/x/stepgain").Text = CStr(gudtSpecData.intMAGICVALGMin)
        xmlDoc.selectSingleNode("/config/MAGICVAL/x/stepoffset").Text = CStr(gudtSpecData.intMAGICVALGMax)
        xmlDoc.selectSingleNode("/config/MAGICVAL/y/stepgain").Text = CStr(gudtSpecData.intMAGICVALOMin)
        xmlDoc.selectSingleNode("/config/MAGICVAL/y/stepoffset").Text = CStr(gudtSpecData.intMAGICVALOMax)
        
        xmlDoc.save gstrXmlPath
    End If
End Sub



Public Property Get CommMode() As CommunicationMode
    CommMode = gudtConfigData.CommMode
End Property

Public Property Let CommMode(enuCommMode As CommunicationMode)
    gudtConfigData.CommMode = enuCommMode
End Property

Public Property Get ComBaud() As String
    ComBaud = gudtConfigData.strComBaud
End Property

Public Property Let ComBaud(strComBaud As String)
    gudtConfigData.strComBaud = strComBaud
End Property

Public Property Get ComID() As Integer
    ComID = gudtConfigData.intComID
End Property

Public Property Let ComID(intComID As Integer)
    gudtConfigData.intComID = intComID
End Property

Public Property Get Model() As String
    Model = gudtConfigData.strModel
End Property

Public Property Let Model(strModel As String)
    gudtConfigData.strModel = strModel
End Property

Public Property Get inputSource() As String
    inputSource = gudtConfigData.strInputSource
End Property

Public Property Let inputSource(strInputSource As String)
    gudtConfigData.strInputSource = strInputSource
End Property

Public Property Get DelayMS() As Long
    DelayMS = gudtConfigData.lngDelayMs
End Property

Public Property Let DelayMS(lngDelayMs As Long)
    gudtConfigData.lngDelayMs = lngDelayMs
End Property

Public Property Get ChannelNum() As Integer
    ChannelNum = gudtConfigData.intChannelNum
End Property

Public Property Let ChannelNum(intChannelNum As Integer)
    gudtConfigData.intChannelNum = intChannelNum
End Property

Public Property Get BarCodeLen() As Integer
    BarCodeLen = gudtConfigData.intBarCodeLen
End Property

Public Property Let BarCodeLen(intBarCodeLen As Integer)
    gudtConfigData.intBarCodeLen = intBarCodeLen
End Property

Public Property Get LvSpec() As Integer
    LvSpec = gudtConfigData.intLvSpec
End Property

Public Property Let LvSpec(intLvSpec As Integer)
    gudtConfigData.intLvSpec = intLvSpec
End Property

Public Property Get VPGModel() As String
    VPGModel = gudtConfigData.strVPGModel
End Property

Public Property Let VPGModel(strVPGModel As String)
    gudtConfigData.strVPGModel = strVPGModel
End Property

Public Property Get VPGTiming() As String
    VPGTiming = gudtConfigData.strVPGTiming
End Property

Public Property Let VPGTiming(strVPGTiming As String)
    gudtConfigData.strVPGTiming = strVPGTiming
End Property

Public Property Get VPG100IRE() As String
    VPG100IRE = gudtConfigData.strVPG100IRE
End Property

Public Property Let VPG100IRE(strVPG100IRE As String)
    gudtConfigData.strVPG100IRE = strVPG100IRE
End Property

Public Property Get VPG80IRE() As String
    VPG80IRE = gudtConfigData.strVPG80IRE
End Property

Public Property Let VPG80IRE(strVPG80IRE As String)
    gudtConfigData.strVPG80IRE = strVPG80IRE
End Property

Public Property Get VPG20IRE() As String
    VPG20IRE = gudtConfigData.strVPG20IRE
End Property

Public Property Let VPG20IRE(strVPG20IRE As String)
    gudtConfigData.strVPG20IRE = strVPG20IRE
End Property

Public Property Get SPECCool1x() As Integer
    SPECCool1x = gudtSpecData.intSPECCool1x
End Property

Public Property Let SPECCool1x(intSPECCool1x As Integer)
    gudtSpecData.intSPECCool1x = intSPECCool1x
End Property

Public Property Get SPECCool1y() As Integer
    SPECCool1y = gudtSpecData.intSPECCool1y
End Property

Public Property Let SPECCool1y(intSPECCool1y As Integer)
    gudtSpecData.intSPECCool1y = intSPECCool1y
End Property

Public Property Get SPECCool1Lv() As Integer
    SPECCool1Lv = gudtSpecData.intSPECCool1Lv
End Property

Public Property Let SPECCool1Lv(intSPECCool1Lv As Integer)
    gudtSpecData.intSPECCool1Lv = intSPECCool1Lv
End Property

Public Property Get SPECNormalx() As Integer
    SPECNormalx = gudtSpecData.intSPECNormalx
End Property

Public Property Let SPECNormalx(intSPECNormalx As Integer)
    gudtSpecData.intSPECNormalx = intSPECNormalx
End Property

Public Property Get SPECNormaly() As Integer
    SPECNormaly = gudtSpecData.intSPECNormaly
End Property

Public Property Let SPECNormaly(intSPECNormaly As Integer)
    gudtSpecData.intSPECNormaly = intSPECNormaly
End Property

Public Property Get SPECNormalLv() As Integer
    SPECNormalLv = gudtSpecData.intSPECNormalLv
End Property

Public Property Let SPECNormalLv(intSPECNormalLv As Integer)
    gudtSpecData.intSPECNormalLv = intSPECNormalLv
End Property

Public Property Get SPECWarm1x() As Integer
    SPECWarm1x = gudtSpecData.intSPECWarm1x
End Property

Public Property Let SPECWarm1x(intSPECWarm1x As Integer)
    gudtSpecData.intSPECWarm1x = intSPECWarm1x
End Property

Public Property Get SPECWarm1y() As Integer
    SPECWarm1y = gudtSpecData.intSPECWarm1y
End Property

Public Property Let SPECWarm1y(intSPECWarm1y As Integer)
    gudtSpecData.intSPECWarm1y = intSPECWarm1y
End Property

Public Property Get SPECWarm1Lv() As Integer
    SPECWarm1Lv = gudtSpecData.intSPECWarm1Lv
End Property

Public Property Let SPECWarm1Lv(intSPECWarm1Lv As Integer)
    gudtSpecData.intSPECWarm1Lv = intSPECWarm1Lv
End Property

Public Property Get TOLCool1xt() As Integer
    TOLCool1xt = gudtSpecData.intTOLCool1xt
End Property

Public Property Let TOLCool1xt(intTOLCool1xt As Integer)
    gudtSpecData.intTOLCool1xt = intTOLCool1xt
End Property

Public Property Get TOLCool1yt() As Integer
    TOLCool1yt = gudtSpecData.intTOLCool1yt
End Property

Public Property Let TOLCool1yt(intTOLCool1yt As Integer)
    gudtSpecData.intTOLCool1yt = intTOLCool1yt
End Property

Public Property Get TOLNormalxt() As Integer
    TOLNormalxt = gudtSpecData.intTOLNormalxt
End Property

Public Property Let TOLNormalxt(intTOLNormalxt As Integer)
    gudtSpecData.intTOLNormalxt = intTOLNormalxt
End Property

Public Property Get TOLNormalyt() As Integer
    TOLNormalyt = gudtSpecData.intTOLNormalyt
End Property

Public Property Let TOLNormalyt(intTOLNormalyt As Integer)
    gudtSpecData.intTOLNormalyt = intTOLNormalyt
End Property

Public Property Get TOLWarm1xt() As Integer
    TOLWarm1xt = gudtSpecData.intTOLWarm1xt
End Property

Public Property Let TOLWarm1xt(intTOLWarm1xt As Integer)
    gudtSpecData.intTOLWarm1xt = intTOLWarm1xt
End Property

Public Property Get TOLWarm1yt() As Integer
    TOLWarm1yt = gudtSpecData.intTOLWarm1yt
End Property

Public Property Let TOLWarm1yt(intTOLWarm1yt As Integer)
    gudtSpecData.intTOLWarm1yt = intTOLWarm1yt
End Property

Public Property Get CHKCool1Cxt() As Integer
    CHKCool1Cxt = gudtSpecData.intCHKCool1Cxt
End Property

Public Property Let CHKCool1Cxt(intCHKCool1Cxt As Integer)
    gudtSpecData.intCHKCool1Cxt = intCHKCool1Cxt
End Property

Public Property Get CHKCool1Cyt() As Integer
    CHKCool1Cyt = gudtSpecData.intCHKCool1Cyt
End Property

Public Property Let CHKCool1Cyt(intCHKCool1Cyt As Integer)
    gudtSpecData.intCHKCool1Cyt = intCHKCool1Cyt
End Property

Public Property Get CHKNormalCxt() As Integer
    CHKNormalCxt = gudtSpecData.intCHKNormalCxt
End Property

Public Property Let CHKNormalCxt(intCHKNormalCxt As Integer)
    gudtSpecData.intCHKNormalCxt = intCHKNormalCxt
End Property

Public Property Get CHKNormalCyt() As Integer
    CHKNormalCyt = gudtSpecData.intCHKNormalCyt
End Property

Public Property Let CHKNormalCyt(intCHKNormalCyt As Integer)
    gudtSpecData.intCHKNormalCyt = intCHKNormalCyt
End Property

Public Property Get CHKWarm1Cxt() As Integer
    CHKWarm1Cxt = gudtSpecData.intCHKWarm1Cxt
End Property

Public Property Let CHKWarm1Cxt(intCHKWarm1Cxt As Integer)
    gudtSpecData.intCHKWarm1Cxt = intCHKWarm1Cxt
End Property

Public Property Get CHKWarm1Cyt() As Integer
    CHKWarm1Cyt = gudtSpecData.intCHKWarm1Cyt
End Property

Public Property Let CHKWarm1Cyt(intCHKWarm1Cyt As Integer)
    gudtSpecData.intCHKWarm1Cyt = intCHKWarm1Cyt
End Property

Public Property Get PRESETGANCool1R() As Integer
    PRESETGANCool1R = gudtSpecData.intPRESETGANCool1R
End Property

Public Property Let PRESETGANCool1R(intPRESETGANCool1R As Integer)
    gudtSpecData.intPRESETGANCool1R = intPRESETGANCool1R
End Property

Public Property Get PRESETGANCool1G() As Integer
    PRESETGANCool1G = gudtSpecData.intPRESETGANCool1G
End Property

Public Property Let PRESETGANCool1G(intPRESETGANCool1G As Integer)
    gudtSpecData.intPRESETGANCool1G = intPRESETGANCool1G
End Property

Public Property Get PRESETGANCool1B() As Integer
    PRESETGANCool1B = gudtSpecData.intPRESETGANCool1B
End Property

Public Property Let PRESETGANCool1B(intPRESETGANCool1B As Integer)
    gudtSpecData.intPRESETGANCool1B = intPRESETGANCool1B
End Property

Public Property Get PRESETGANNormalR() As Integer
    PRESETGANNormalR = gudtSpecData.intPRESETGANNormalR
End Property

Public Property Let PRESETGANNormalR(intPRESETGANNormalR As Integer)
    gudtSpecData.intPRESETGANNormalR = intPRESETGANNormalR
End Property

Public Property Get PRESETGANNormalG() As Integer
    PRESETGANNormalG = gudtSpecData.intPRESETGANNormalG
End Property

Public Property Let PRESETGANNormalG(intPRESETGANNormalG As Integer)
    gudtSpecData.intPRESETGANNormalG = intPRESETGANNormalG
End Property

Public Property Get PRESETGANNormalB() As Integer
    PRESETGANNormalB = gudtSpecData.intPRESETGANNormalB
End Property

Public Property Let PRESETGANNormalB(intPRESETGANNormalB As Integer)
    gudtSpecData.intPRESETGANNormalB = intPRESETGANNormalB
End Property

Public Property Get PRESETGANWarm1R() As Integer
    PRESETGANWarm1R = gudtSpecData.intPRESETGANWarm1R
End Property

Public Property Let PRESETGANWarm1R(intPRESETGANWarm1R As Integer)
    gudtSpecData.intPRESETGANWarm1R = intPRESETGANWarm1R
End Property

Public Property Get PRESETGANWarm1G() As Integer
    PRESETGANWarm1G = gudtSpecData.intPRESETGANWarm1G
End Property

Public Property Let PRESETGANWarm1G(intPRESETGANWarm1G As Integer)
    gudtSpecData.intPRESETGANWarm1G = intPRESETGANWarm1G
End Property

Public Property Get PRESETGANWarm1B() As Integer
    PRESETGANWarm1B = gudtSpecData.intPRESETGANWarm1B
End Property

Public Property Let PRESETGANWarm1B(intPRESETGANWarm1B As Integer)
    gudtSpecData.intPRESETGANWarm1B = intPRESETGANWarm1B
End Property

Public Property Get PRESETOFFCool1R() As Integer
    PRESETOFFCool1R = gudtSpecData.intPRESETOFFCool1R
End Property

Public Property Let PRESETOFFCool1R(intPRESETOFFCool1R As Integer)
    gudtSpecData.intPRESETOFFCool1R = intPRESETOFFCool1R
End Property

Public Property Get PRESETOFFCool1G() As Integer
    PRESETOFFCool1G = gudtSpecData.intPRESETOFFCool1G
End Property

Public Property Let PRESETOFFCool1G(intPRESETOFFCool1G As Integer)
    gudtSpecData.intPRESETOFFCool1G = intPRESETOFFCool1G
End Property

Public Property Get PRESETOFFCool1B() As Integer
    PRESETOFFCool1B = gudtSpecData.intPRESETOFFCool1B
End Property

Public Property Let PRESETOFFCool1B(intPRESETOFFCool1B As Integer)
    gudtSpecData.intPRESETOFFCool1B = intPRESETOFFCool1B
End Property

Public Property Get PRESETOFFNormalR() As Integer
    PRESETOFFNormalR = gudtSpecData.intPRESETOFFNormalR
End Property

Public Property Let PRESETOFFNormalR(intPRESETOFFNormalR As Integer)
    gudtSpecData.intPRESETOFFNormalR = intPRESETOFFNormalR
End Property

Public Property Get PRESETOFFNormalG() As Integer
    PRESETOFFNormalG = gudtSpecData.intPRESETOFFNormalG
End Property

Public Property Let PRESETOFFNormalG(intPRESETOFFNormalG As Integer)
    gudtSpecData.intPRESETOFFNormalG = intPRESETOFFNormalG
End Property

Public Property Get PRESETOFFNormalB() As Integer
    PRESETOFFNormalB = gudtSpecData.intPRESETOFFNormalB
End Property

Public Property Let PRESETOFFNormalB(intPRESETOFFNormalB As Integer)
    gudtSpecData.intPRESETOFFNormalB = intPRESETOFFNormalB
End Property

Public Property Get PRESETOFFWarm1R() As Integer
    PRESETOFFWarm1R = gudtSpecData.intPRESETOFFWarm1R
End Property

Public Property Let PRESETOFFWarm1R(intPRESETOFFWarm1R As Integer)
    gudtSpecData.intPRESETOFFWarm1R = intPRESETOFFWarm1R
End Property

Public Property Get PRESETOFFWarm1G() As Integer
    PRESETOFFWarm1G = gudtSpecData.intPRESETOFFWarm1G
End Property

Public Property Let PRESETOFFWarm1G(intPRESETOFFWarm1G As Integer)
    gudtSpecData.intPRESETOFFWarm1G = intPRESETOFFWarm1G
End Property

Public Property Get PRESETOFFWarm1B() As Integer
    PRESETOFFWarm1B = gudtSpecData.intPRESETOFFWarm1B
End Property

Public Property Let PRESETOFFWarm1B(intPRESETOFFWarm1B As Integer)
    gudtSpecData.intPRESETOFFWarm1B = intPRESETOFFWarm1B
End Property

Public Property Get CLEVELRGBGMin() As Integer
    CLEVELRGBGMin = gudtSpecData.intCLEVELRGBGMin
End Property

Public Property Let CLEVELRGBGMin(intCLEVELRGBGMin As Integer)
    gudtSpecData.intCLEVELRGBGMin = intCLEVELRGBGMin
End Property

Public Property Get CLEVELRGBGMax() As Integer
    CLEVELRGBGMax = gudtSpecData.intCLEVELRGBGMax
End Property

Public Property Let CLEVELRGBGMax(intCLEVELRGBGMax As Integer)
    gudtSpecData.intCLEVELRGBGMax = intCLEVELRGBGMax
End Property

Public Property Get CLEVELRGBOMin() As Integer
    CLEVELRGBOMin = gudtSpecData.intCLEVELRGBOMin
End Property

Public Property Let CLEVELRGBOMin(intCLEVELRGBOMin As Integer)
    gudtSpecData.intCLEVELRGBOMin = intCLEVELRGBOMin
End Property

Public Property Get CLEVELRGBOMax() As Integer
    CLEVELRGBOMax = gudtSpecData.intCLEVELRGBOMax
End Property

Public Property Let CLEVELRGBOMax(intCLEVELRGBOMax As Integer)
    gudtSpecData.intCLEVELRGBOMax = intCLEVELRGBOMax
End Property

Public Property Get MAGICVALGMin() As Integer
    MAGICVALGMin = gudtSpecData.intMAGICVALGMin
End Property

Public Property Let MAGICVALGMin(intMAGICVALGMin As Integer)
    gudtSpecData.intMAGICVALGMin = intMAGICVALGMin
End Property

Public Property Get MAGICVALGMax() As Integer
    MAGICVALGMax = gudtSpecData.intMAGICVALGMax
End Property

Public Property Let MAGICVALGMax(intMAGICVALGMax As Integer)
    gudtSpecData.intMAGICVALGMax = intMAGICVALGMax
End Property

Public Property Get MAGICVALOMin() As Integer
    MAGICVALOMin = gudtSpecData.intMAGICVALOMin
End Property

Public Property Let MAGICVALOMin(intMAGICVALOMin As Integer)
    gudtSpecData.intMAGICVALOMin = intMAGICVALOMin
End Property

Public Property Get MAGICVALOMax() As Integer
    MAGICVALOMax = gudtSpecData.intMAGICVALOMax
End Property

Public Property Let MAGICVALOMax(intMAGICVALOMax As Integer)
    gudtSpecData.intMAGICVALOMax = intMAGICVALOMax
End Property


Public Property Get I2cClockRate() As String
    I2cClockRate = gudtConfigData.lngI2cClockRate
End Property

Public Property Let I2cClockRate(lngI2cClockRate As String)
    gudtConfigData.lngI2cClockRate = lngI2cClockRate
End Property

Public Property Get EnableCool2() As Boolean
    EnableCool2 = gudtConfigData.bolEnableCool2
End Property

Public Property Let EnableCool2(bolEnableCool2 As Boolean)
    gudtConfigData.bolEnableCool2 = bolEnableCool2
End Property

Public Property Get EnableCool1() As Boolean
    EnableCool1 = gudtConfigData.bolEnableCool1
End Property

Public Property Let EnableCool1(bolEnableCool1 As Boolean)
    gudtConfigData.bolEnableCool1 = bolEnableCool1
End Property

Public Property Get EnableNormal() As Boolean
    EnableNormal = gudtConfigData.bolEnableNormal
End Property

Public Property Let EnableNormal(bolEnableNormal As Boolean)
    gudtConfigData.bolEnableNormal = bolEnableNormal
End Property

Public Property Get EnableWarm1() As Boolean
    EnableWarm1 = gudtConfigData.bolEnableWarm1
End Property

Public Property Let EnableWarm1(bolEnableWarm1 As Boolean)
    gudtConfigData.bolEnableWarm1 = bolEnableWarm1
End Property

Public Property Get EnableWarm2() As Boolean
    EnableWarm2 = gudtConfigData.bolEnableWarm2
End Property

Public Property Let EnableWarm2(bolEnableWarm2 As Boolean)
    gudtConfigData.bolEnableWarm2 = bolEnableWarm2
End Property

Public Property Get EnableChkColor() As Boolean
    EnableChkColor = gudtConfigData.bolEnableChkColor
End Property

Public Property Let EnableChkColor(bolEnableChkColor As Boolean)
    gudtConfigData.bolEnableChkColor = bolEnableChkColor
End Property

Public Property Get EnableAdjOffset() As Boolean
    EnableAdjOffset = gudtConfigData.bolEnableAdjOffset
End Property

Public Property Let EnableAdjOffset(bolEnableAdjOffset As Boolean)
    gudtConfigData.bolEnableAdjOffset = bolEnableAdjOffset
End Property

Public Property Get ChipSet() As String
    ChipSet = gudtConfigData.strChipSet
End Property

