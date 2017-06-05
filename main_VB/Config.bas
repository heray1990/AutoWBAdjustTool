Attribute VB_Name = "Config"
'**********************************************
' Class module for handling config.xml of the
' application.
'**********************************************

Option Explicit
Public gstrXmlPath As String
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long




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

Public Sub LoadConfigData1()

    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(gstrXmlPath)
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        rConfigData.intSPECCool1x = val(xmlDoc.selectSingleNode("/config/SPEC/cool1/x").Text)
        rConfigData.intSPECCool1y = val(xmlDoc.selectSingleNode("/config/SPEC/cool1/y").Text)
        rConfigData.intSPECCool1Lv = val(xmlDoc.selectSingleNode("/config/SPEC/cool1/Lv").Text)
        rConfigData.intSPECNormalx = val(xmlDoc.selectSingleNode("/config/SPEC/normal/x").Text)
        rConfigData.intSPECNormaly = val(xmlDoc.selectSingleNode("/config/SPEC/normal/y").Text)
        rConfigData.intSPECNormalLv = val(xmlDoc.selectSingleNode("/config/SPEC/normal/Lv").Text)
        rConfigData.intSPECWarm1x = val(xmlDoc.selectSingleNode("/config/SPEC/warm1/x").Text)
        rConfigData.intSPECWarm1y = val(xmlDoc.selectSingleNode("/config/SPEC/warm1/y").Text)
        rConfigData.intSPECWarm1Lv = val(xmlDoc.selectSingleNode("/config/SPEC/warm1/Lv").Text)
        rConfigData.intTOLCool1xt = val(xmlDoc.selectSingleNode("/config/TOL/cool1/xt").Text)
        rConfigData.intTOLCool1yt = val(xmlDoc.selectSingleNode("/config/TOL/cool1/yt").Text)
        rConfigData.intTOLNormalxt = val(xmlDoc.selectSingleNode("/config/TOL/normal/xt").Text)
        rConfigData.intTOLNormalyt = val(xmlDoc.selectSingleNode("/config/TOL/normal/yt").Text)
        rConfigData.intTOLWarm1xt = val(xmlDoc.selectSingleNode("/config/TOL/warm1/xt").Text)
        rConfigData.intTOLWarm1yt = val(xmlDoc.selectSingleNode("/config/TOL/warm1/yt").Text)
        rConfigData.intCHKCool1Cxt = val(xmlDoc.selectSingleNode("/config/CHK/cool1/cxt").Text)
        rConfigData.intCHKCool1Cyt = val(xmlDoc.selectSingleNode("/config/CHK/cool1/cyt").Text)
        rConfigData.intCHKNormalCxt = val(xmlDoc.selectSingleNode("/config/CHK/normal/cxt").Text)
        rConfigData.intCHKNormalCyt = val(xmlDoc.selectSingleNode("/config/CHK/normal/cyt").Text)
        rConfigData.intCHKWarm1Cxt = val(xmlDoc.selectSingleNode("/config/CHK/warm1/cxt").Text)
        rConfigData.intCHKWarm1Cyt = val(xmlDoc.selectSingleNode("/config/CHK/warm1/cyt").Text)
        rConfigData.intPRESETGANCool1R = val(xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/R").Text)
        rConfigData.intPRESETGANCool1G = val(xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/G").Text)
        rConfigData.intPRESETGANCool1B = val(xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/B").Text)
        rConfigData.intPRESETGANNormalR = val(xmlDoc.selectSingleNode("/config/PRESETGAN/normal/R").Text)
        rConfigData.intPRESETGANNormalG = val(xmlDoc.selectSingleNode("/config/PRESETGAN/normal/G").Text)
        rConfigData.intPRESETGANNormalB = val(xmlDoc.selectSingleNode("/config/PRESETGAN/normal/B").Text)
        rConfigData.intPRESETGANWarm1R = val(xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/R").Text)
        rConfigData.intPRESETGANWarm1G = val(xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/G").Text)
        rConfigData.intPRESETGANWarm1B = val(xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/B").Text)
        rConfigData.intPRESETOFFCool1R = val(xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/R").Text)
        rConfigData.intPRESETOFFCool1G = val(xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/G").Text)
        rConfigData.intPRESETOFFCool1B = val(xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/B").Text)
        rConfigData.intPRESETOFFNormalR = val(xmlDoc.selectSingleNode("/config/PRESETOFF/normal/R").Text)
        rConfigData.intPRESETOFFNormalG = val(xmlDoc.selectSingleNode("/config/PRESETOFF/normal/G").Text)
        rConfigData.intPRESETOFFNormalB = val(xmlDoc.selectSingleNode("/config/PRESETOFF/normal/B").Text)
        rConfigData.intPRESETOFFWarm1R = val(xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/R").Text)
        rConfigData.intPRESETOFFWarm1G = val(xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/G").Text)
        rConfigData.intPRESETOFFWarm1B = val(xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/B").Text)
        rConfigData.intCLEVELRGBGMin = val(xmlDoc.selectSingleNode("/config/CLEVELRGB/gain/min").Text)
        rConfigData.intCLEVELRGBGMax = val(xmlDoc.selectSingleNode("/config/CLEVELRGB/gain/max").Text)
        rConfigData.intCLEVELRGBOMin = val(xmlDoc.selectSingleNode("/config/CLEVELRGB/offset/min").Text)
        rConfigData.intCLEVELRGBOMax = val(xmlDoc.selectSingleNode("/config/CLEVELRGB/offset/max").Text)
        rConfigData.intMAGICVALGMin = val(xmlDoc.selectSingleNode("/config/MAGICVAL/x/stepgain").Text)
        rConfigData.intMAGICVALGMax = val(xmlDoc.selectSingleNode("/config/MAGICVAL/x/stepoffset").Text)
        rConfigData.intMAGICVALOMin = val(xmlDoc.selectSingleNode("/config/MAGICVAL/y/stepgain").Text)
        rConfigData.intMAGICVALOMax = val(xmlDoc.selectSingleNode("/config/MAGICVAL/y/stepoffset").Text)
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

Public Sub SaveConfigData1()
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(gstrXmlPath)
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        xmlDoc.selectSingleNode("/config/SPEC/cool1/x").Text = CStr(rConfigData.intSPECCool1x)
        xmlDoc.selectSingleNode("/config/SPEC/cool1/y").Text = CStr(rConfigData.intSPECCool1y)
        xmlDoc.selectSingleNode("/config/SPEC/cool1/Lv").Text = CStr(rConfigData.intSPECCool1Lv)
        xmlDoc.selectSingleNode("/config/SPEC/normal/x").Text = CStr(rConfigData.intSPECNormalx)
        xmlDoc.selectSingleNode("/config/SPEC/normal/y").Text = CStr(rConfigData.intSPECNormaly)
        xmlDoc.selectSingleNode("/config/SPEC/normal/Lv").Text = CStr(rConfigData.intSPECNormalLv)
        xmlDoc.selectSingleNode("/config/SPEC/warm1/x").Text = CStr(rConfigData.intSPECWarm1x)
        xmlDoc.selectSingleNode("/config/SPEC/warm1/y").Text = CStr(rConfigData.intSPECWarm1y)
        xmlDoc.selectSingleNode("/config/SPEC/warm1/Lv").Text = CStr(rConfigData.intSPECWarm1Lv)
        xmlDoc.selectSingleNode("/config/TOL/cool1/xt").Text = CStr(rConfigData.intTOLCool1xt)
        xmlDoc.selectSingleNode("/config/TOL/cool1/yt").Text = CStr(rConfigData.intTOLCool1yt)
        xmlDoc.selectSingleNode("/config/TOL/normal/xt").Text = CStr(rConfigData.intTOLNormalxt)
        xmlDoc.selectSingleNode("/config/TOL/normal/yt").Text = CStr(rConfigData.intTOLNormalyt)
        xmlDoc.selectSingleNode("/config/TOL/warm1/xt").Text = CStr(rConfigData.intTOLWarm1xt)
        xmlDoc.selectSingleNode("/config/TOL/warm1/yt").Text = CStr(rConfigData.intTOLWarm1yt)
        xmlDoc.selectSingleNode("/config/CHK/cool1/cxt").Text = CStr(rConfigData.intCHKCool1Cxt)
        xmlDoc.selectSingleNode("/config/CHK/cool1/cyt").Text = CStr(rConfigData.intCHKCool1Cyt)
        xmlDoc.selectSingleNode("/config/CHK/normal/cxt").Text = CStr(rConfigData.intCHKNormalCxt)
        xmlDoc.selectSingleNode("/config/CHK/normal/cyt").Text = CStr(rConfigData.intCHKNormalCyt)
        xmlDoc.selectSingleNode("/config/CHK/warm1/cxt").Text = CStr(rConfigData.intCHKWarm1Cxt)
        xmlDoc.selectSingleNode("/config/CHK/warm1/cyt").Text = CStr(rConfigData.intCHKWarm1Cyt)
        xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/R").Text = CStr(rConfigData.intPRESETGANCool1R)
        xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/G").Text = CStr(rConfigData.intPRESETGANCool1G)
        xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/B").Text = CStr(rConfigData.intPRESETGANCool1B)
        xmlDoc.selectSingleNode("/config/PRESETGAN/normal/R").Text = CStr(rConfigData.intPRESETGANNormalR)
        xmlDoc.selectSingleNode("/config/PRESETGAN/normal/G").Text = CStr(rConfigData.intPRESETGANNormalG)
        xmlDoc.selectSingleNode("/config/PRESETGAN/normal/B").Text = CStr(rConfigData.intPRESETGANNormalB)
        xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/R").Text = CStr(rConfigData.intPRESETGANWarm1R)
        xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/G").Text = CStr(rConfigData.intPRESETGANWarm1G)
        xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/B").Text = CStr(rConfigData.intPRESETGANWarm1B)
        xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/R").Text = CStr(rConfigData.intPRESETOFFCool1R)
        xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/G").Text = CStr(rConfigData.intPRESETOFFCool1G)
        xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/B").Text = CStr(rConfigData.intPRESETOFFCool1B)
        xmlDoc.selectSingleNode("/config/PRESETOFF/normal/R").Text = CStr(rConfigData.intPRESETOFFNormalR)
        xmlDoc.selectSingleNode("/config/PRESETOFF/normal/G").Text = CStr(rConfigData.intPRESETOFFNormalG)
        xmlDoc.selectSingleNode("/config/PRESETOFF/normal/B").Text = CStr(rConfigData.intPRESETOFFNormalB)
        xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/R").Text = CStr(rConfigData.intPRESETOFFWarm1R)
        xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/G").Text = CStr(rConfigData.intPRESETOFFWarm1G)
        xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/B").Text = CStr(rConfigData.intPRESETOFFWarm1B)
        xmlDoc.selectSingleNode("/config/CLEVELRGB/gain/min").Text = CStr(rConfigData.intCLEVELRGBGMin)
        xmlDoc.selectSingleNode("/config/CLEVELRGB/gain/max").Text = CStr(rConfigData.intCLEVELRGBGMax)
        xmlDoc.selectSingleNode("/config/CLEVELRGB/offset/min").Text = CStr(rConfigData.intCLEVELRGBOMin)
        xmlDoc.selectSingleNode("/config/CLEVELRGB/offset/max").Text = CStr(rConfigData.intCLEVELRGBOMax)
        xmlDoc.selectSingleNode("/config/MAGICVAL/x/stepgain").Text = CStr(rConfigData.intMAGICVALGMin)
        xmlDoc.selectSingleNode("/config/MAGICVAL/x/stepoffset").Text = CStr(rConfigData.intMAGICVALGMax)
        xmlDoc.selectSingleNode("/config/MAGICVAL/y/stepgain").Text = CStr(rConfigData.intMAGICVALOMin)
        xmlDoc.selectSingleNode("/config/MAGICVAL/y/stepoffset").Text = CStr(rConfigData.intMAGICVALOMax)
        
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
    SPECCool1x = rConfigData.intSPECCool1x
End Property

Public Property Let SPECCool1x(intSPECCool1x As Integer)
    rConfigData.intSPECCool1x = intSPECCool1x
End Property

Public Property Get SPECCool1y() As Integer
    SPECCool1y = rConfigData.intSPECCool1y
End Property

Public Property Let SPECCool1y(intSPECCool1y As Integer)
    rConfigData.intSPECCool1y = intSPECCool1y
End Property

Public Property Get SPECCool1Lv() As Integer
    SPECCool1Lv = rConfigData.intSPECCool1Lv
End Property

Public Property Let SPECCool1Lv(intSPECCool1Lv As Integer)
    rConfigData.intSPECCool1Lv = intSPECCool1Lv
End Property

Public Property Get SPECNormalx() As Integer
    SPECNormalx = rConfigData.intSPECNormalx
End Property

Public Property Let SPECNormalx(intSPECNormalx As Integer)
    rConfigData.intSPECNormalx = intSPECNormalx
End Property

Public Property Get SPECNormaly() As Integer
    SPECNormaly = rConfigData.intSPECNormaly
End Property

Public Property Let SPECNormaly(intSPECNormaly As Integer)
    rConfigData.intSPECNormaly = intSPECNormaly
End Property

Public Property Get SPECNormalLv() As Integer
    SPECNormalLv = rConfigData.intSPECNormalLv
End Property

Public Property Let SPECNormalLv(intSPECNormalLv As Integer)
    rConfigData.intSPECNormalLv = intSPECNormalLv
End Property

Public Property Get SPECWarm1x() As Integer
    SPECWarm1x = rConfigData.intSPECWarm1x
End Property

Public Property Let SPECWarm1x(intSPECWarm1x As Integer)
    rConfigData.intSPECWarm1x = intSPECWarm1x
End Property

Public Property Get SPECWarm1y() As Integer
    SPECWarm1y = rConfigData.intSPECWarm1y
End Property

Public Property Let SPECWarm1y(intSPECWarm1y As Integer)
    rConfigData.intSPECWarm1y = intSPECWarm1y
End Property

Public Property Get SPECWarm1Lv() As Integer
    SPECWarm1Lv = rConfigData.intSPECWarm1Lv
End Property

Public Property Let SPECWarm1Lv(intSPECWarm1Lv As Integer)
    rConfigData.intSPECWarm1Lv = intSPECWarm1Lv
End Property

Public Property Get TOLCool1xt() As Integer
    TOLCool1xt = rConfigData.intTOLCool1xt
End Property

Public Property Let TOLCool1xt(intTOLCool1xt As Integer)
    rConfigData.intTOLCool1xt = intTOLCool1xt
End Property

Public Property Get TOLCool1yt() As Integer
    TOLCool1yt = rConfigData.intTOLCool1yt
End Property

Public Property Let TOLCool1yt(intTOLCool1yt As Integer)
    rConfigData.intTOLCool1yt = intTOLCool1yt
End Property

Public Property Get TOLNormalxt() As Integer
    TOLNormalxt = rConfigData.intTOLNormalxt
End Property

Public Property Let TOLNormalxt(intTOLNormalxt As Integer)
    rConfigData.intTOLNormalxt = intTOLNormalxt
End Property

Public Property Get TOLNormalyt() As Integer
    TOLNormalyt = rConfigData.intTOLNormalyt
End Property

Public Property Let TOLNormalyt(intTOLNormalyt As Integer)
    rConfigData.intTOLNormalyt = intTOLNormalyt
End Property

Public Property Get TOLWarm1xt() As Integer
    TOLWarm1xt = rConfigData.intTOLWarm1xt
End Property

Public Property Let TOLWarm1xt(intTOLWarm1xt As Integer)
    rConfigData.intTOLWarm1xt = intTOLWarm1xt
End Property

Public Property Get TOLWarm1yt() As Integer
    TOLWarm1yt = rConfigData.intTOLWarm1yt
End Property

Public Property Let TOLWarm1yt(intTOLWarm1yt As Integer)
    rConfigData.intTOLWarm1yt = intTOLWarm1yt
End Property

Public Property Get CHKCool1Cxt() As Integer
    CHKCool1Cxt = rConfigData.intCHKCool1Cxt
End Property

Public Property Let CHKCool1Cxt(intCHKCool1Cxt As Integer)
    rConfigData.intCHKCool1Cxt = intCHKCool1Cxt
End Property

Public Property Get CHKCool1Cyt() As Integer
    CHKCool1Cyt = rConfigData.intCHKCool1Cyt
End Property

Public Property Let CHKCool1Cyt(intCHKCool1Cyt As Integer)
    rConfigData.intCHKCool1Cyt = intCHKCool1Cyt
End Property

Public Property Get CHKNormalCxt() As Integer
    CHKNormalCxt = rConfigData.intCHKNormalCxt
End Property

Public Property Let CHKNormalCxt(intCHKNormalCxt As Integer)
    rConfigData.intCHKNormalCxt = intCHKNormalCxt
End Property

Public Property Get CHKNormalCyt() As Integer
    CHKNormalCyt = rConfigData.intCHKNormalCyt
End Property

Public Property Let CHKNormalCyt(intCHKNormalCyt As Integer)
    rConfigData.intCHKNormalCyt = intCHKNormalCyt
End Property

Public Property Get CHKWarm1Cxt() As Integer
    CHKWarm1Cxt = rConfigData.intCHKWarm1Cxt
End Property

Public Property Let CHKWarm1Cxt(intCHKWarm1Cxt As Integer)
    rConfigData.intCHKWarm1Cxt = intCHKWarm1Cxt
End Property

Public Property Get CHKWarm1Cyt() As Integer
    CHKWarm1Cyt = rConfigData.intCHKWarm1Cyt
End Property

Public Property Let CHKWarm1Cyt(intCHKWarm1Cyt As Integer)
    rConfigData.intCHKWarm1Cyt = intCHKWarm1Cyt
End Property

Public Property Get PRESETGANCool1R() As Integer
    PRESETGANCool1R = rConfigData.intPRESETGANCool1R
End Property

Public Property Let PRESETGANCool1R(intPRESETGANCool1R As Integer)
    rConfigData.intPRESETGANCool1R = intPRESETGANCool1R
End Property

Public Property Get PRESETGANCool1G() As Integer
    PRESETGANCool1G = rConfigData.intPRESETGANCool1G
End Property

Public Property Let PRESETGANCool1G(intPRESETGANCool1G As Integer)
    rConfigData.intPRESETGANCool1G = intPRESETGANCool1G
End Property

Public Property Get PRESETGANCool1B() As Integer
    PRESETGANCool1B = rConfigData.intPRESETGANCool1B
End Property

Public Property Let PRESETGANCool1B(intPRESETGANCool1B As Integer)
    rConfigData.intPRESETGANCool1B = intPRESETGANCool1B
End Property

Public Property Get PRESETGANNormalR() As Integer
    PRESETGANNormalR = rConfigData.intPRESETGANNormalR
End Property

Public Property Let PRESETGANNormalR(intPRESETGANNormalR As Integer)
    rConfigData.intPRESETGANNormalR = intPRESETGANNormalR
End Property

Public Property Get PRESETGANNormalG() As Integer
    PRESETGANNormalG = rConfigData.intPRESETGANNormalG
End Property

Public Property Let PRESETGANNormalG(intPRESETGANNormalG As Integer)
    rConfigData.intPRESETGANNormalG = intPRESETGANNormalG
End Property

Public Property Get PRESETGANNormalB() As Integer
    PRESETGANNormalB = rConfigData.intPRESETGANNormalB
End Property

Public Property Let PRESETGANNormalB(intPRESETGANNormalB As Integer)
    rConfigData.intPRESETGANNormalB = intPRESETGANNormalB
End Property

Public Property Get PRESETGANWarm1R() As Integer
    PRESETGANWarm1R = rConfigData.intPRESETGANWarm1R
End Property

Public Property Let PRESETGANWarm1R(intPRESETGANWarm1R As Integer)
    rConfigData.intPRESETGANWarm1R = intPRESETGANWarm1R
End Property

Public Property Get PRESETGANWarm1G() As Integer
    PRESETGANWarm1G = rConfigData.intPRESETGANWarm1G
End Property

Public Property Let PRESETGANWarm1G(intPRESETGANWarm1G As Integer)
    rConfigData.intPRESETGANWarm1G = intPRESETGANWarm1G
End Property

Public Property Get PRESETGANWarm1B() As Integer
    PRESETGANWarm1B = rConfigData.intPRESETGANWarm1B
End Property

Public Property Let PRESETGANWarm1B(intPRESETGANWarm1B As Integer)
    rConfigData.intPRESETGANWarm1B = intPRESETGANWarm1B
End Property

Public Property Get PRESETOFFCool1R() As Integer
    PRESETOFFCool1R = rConfigData.intPRESETOFFCool1R
End Property

Public Property Let PRESETOFFCool1R(intPRESETOFFCool1R As Integer)
    rConfigData.intPRESETOFFCool1R = intPRESETOFFCool1R
End Property

Public Property Get PRESETOFFCool1G() As Integer
    PRESETOFFCool1G = rConfigData.intPRESETOFFCool1G
End Property

Public Property Let PRESETOFFCool1G(intPRESETOFFCool1G As Integer)
    rConfigData.intPRESETOFFCool1G = intPRESETOFFCool1G
End Property

Public Property Get PRESETOFFCool1B() As Integer
    PRESETOFFCool1B = rConfigData.intPRESETOFFCool1B
End Property

Public Property Let PRESETOFFCool1B(intPRESETOFFCool1B As Integer)
    rConfigData.intPRESETOFFCool1B = intPRESETOFFCool1B
End Property

Public Property Get PRESETOFFNormalR() As Integer
    PRESETOFFNormalR = rConfigData.intPRESETOFFNormalR
End Property

Public Property Let PRESETOFFNormalR(intPRESETOFFNormalR As Integer)
    rConfigData.intPRESETOFFNormalR = intPRESETOFFNormalR
End Property

Public Property Get PRESETOFFNormalG() As Integer
    PRESETOFFNormalG = rConfigData.intPRESETOFFNormalG
End Property

Public Property Let PRESETOFFNormalG(intPRESETOFFNormalG As Integer)
    rConfigData.intPRESETOFFNormalG = intPRESETOFFNormalG
End Property

Public Property Get PRESETOFFNormalB() As Integer
    PRESETOFFNormalB = rConfigData.intPRESETOFFNormalB
End Property

Public Property Let PRESETOFFNormalB(intPRESETOFFNormalB As Integer)
    rConfigData.intPRESETOFFNormalB = intPRESETOFFNormalB
End Property

Public Property Get PRESETOFFWarm1R() As Integer
    PRESETOFFWarm1R = rConfigData.intPRESETOFFWarm1R
End Property

Public Property Let PRESETOFFWarm1R(intPRESETOFFWarm1R As Integer)
    rConfigData.intPRESETOFFWarm1R = intPRESETOFFWarm1R
End Property

Public Property Get PRESETOFFWarm1G() As Integer
    PRESETOFFWarm1G = rConfigData.intPRESETOFFWarm1G
End Property

Public Property Let PRESETOFFWarm1G(intPRESETOFFWarm1G As Integer)
    rConfigData.intPRESETOFFWarm1G = intPRESETOFFWarm1G
End Property

Public Property Get PRESETOFFWarm1B() As Integer
    PRESETOFFWarm1B = rConfigData.intPRESETOFFWarm1B
End Property

Public Property Let PRESETOFFWarm1B(intPRESETOFFWarm1B As Integer)
    rConfigData.intPRESETOFFWarm1B = intPRESETOFFWarm1B
End Property

Public Property Get CLEVELRGBGMin() As Integer
    CLEVELRGBGMin = rConfigData.intCLEVELRGBGMin
End Property

Public Property Let CLEVELRGBGMin(intCLEVELRGBGMin As Integer)
    rConfigData.intCLEVELRGBGMin = intCLEVELRGBGMin
End Property

Public Property Get CLEVELRGBGMax() As Integer
    CLEVELRGBGMax = rConfigData.intCLEVELRGBGMax
End Property

Public Property Let CLEVELRGBGMax(intCLEVELRGBGMax As Integer)
    rConfigData.intCLEVELRGBGMax = intCLEVELRGBGMax
End Property

Public Property Get CLEVELRGBOMin() As Integer
    CLEVELRGBOMin = rConfigData.intCLEVELRGBOMin
End Property

Public Property Let CLEVELRGBOMin(intCLEVELRGBOMin As Integer)
    rConfigData.intCLEVELRGBOMin = intCLEVELRGBOMin
End Property

Public Property Get CLEVELRGBOMax() As Integer
    CLEVELRGBOMax = rConfigData.intCLEVELRGBOMax
End Property

Public Property Let CLEVELRGBOMax(intCLEVELRGBOMax As Integer)
    rConfigData.intCLEVELRGBOMax = intCLEVELRGBOMax
End Property

Public Property Get MAGICVALGMin() As Integer
    MAGICVALGMin = rConfigData.intMAGICVALGMin
End Property

Public Property Let MAGICVALGMin(intMAGICVALGMin As Integer)
    rConfigData.intMAGICVALGMin = intMAGICVALGMin
End Property

Public Property Get MAGICVALGMax() As Integer
    MAGICVALGMax = rConfigData.intMAGICVALGMax
End Property

Public Property Let MAGICVALGMax(intMAGICVALGMax As Integer)
    rConfigData.intMAGICVALGMax = intMAGICVALGMax
End Property

Public Property Get MAGICVALOMin() As Integer
    MAGICVALOMin = rConfigData.intMAGICVALOMin
End Property

Public Property Let MAGICVALOMin(intMAGICVALOMin As Integer)
    rConfigData.intMAGICVALOMin = intMAGICVALOMin
End Property

Public Property Get MAGICVALOMax() As Integer
    MAGICVALOMax = rConfigData.intMAGICVALOMax
End Property

Public Property Let MAGICVALOMax(intMAGICVALOMax As Integer)
    rConfigData.intMAGICVALOMax = intMAGICVALOMax
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

