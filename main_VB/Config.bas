Attribute VB_Name = "Config"
'**********************************************
' Class module for handling config.xml of the
' application.
'**********************************************

Option Explicit
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long


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

Public Sub SaveSpecData(strColorTemp As String)
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(gstrXmlPath)
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        If strColorTemp = COLORTEMP_COOL1 Then
            xmlDoc.selectSingleNode("/config/SPEC/cool1/x").Text = CStr(gudtSpecData.intSPECCool1x)
            xmlDoc.selectSingleNode("/config/SPEC/cool1/y").Text = CStr(gudtSpecData.intSPECCool1y)
            xmlDoc.selectSingleNode("/config/TOL/cool1/xt").Text = CStr(gudtSpecData.intTOLCool1xt)
            xmlDoc.selectSingleNode("/config/TOL/cool1/yt").Text = CStr(gudtSpecData.intTOLCool1yt)
            xmlDoc.selectSingleNode("/config/CHK/cool1/cxt").Text = CStr(gudtSpecData.intCHKCool1Cxt)
            xmlDoc.selectSingleNode("/config/CHK/cool1/cyt").Text = CStr(gudtSpecData.intCHKCool1Cyt)
            xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/R").Text = CStr(gudtSpecData.intPRESETGANCool1R)
            xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/G").Text = CStr(gudtSpecData.intPRESETGANCool1G)
            xmlDoc.selectSingleNode("/config/PRESETGAN/cool1/B").Text = CStr(gudtSpecData.intPRESETGANCool1B)
            xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/R").Text = CStr(gudtSpecData.intPRESETOFFCool1R)
            xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/G").Text = CStr(gudtSpecData.intPRESETOFFCool1G)
            xmlDoc.selectSingleNode("/config/PRESETOFF/cool1/B").Text = CStr(gudtSpecData.intPRESETOFFCool1B)
        ElseIf strColorTemp = COLORTEMP_STANDARD Then
            xmlDoc.selectSingleNode("/config/SPEC/normal/x").Text = CStr(gudtSpecData.intSPECNormalx)
            xmlDoc.selectSingleNode("/config/SPEC/normal/y").Text = CStr(gudtSpecData.intSPECNormaly)
            xmlDoc.selectSingleNode("/config/TOL/normal/xt").Text = CStr(gudtSpecData.intTOLNormalxt)
            xmlDoc.selectSingleNode("/config/TOL/normal/yt").Text = CStr(gudtSpecData.intTOLNormalyt)
            xmlDoc.selectSingleNode("/config/CHK/normal/cxt").Text = CStr(gudtSpecData.intCHKNormalCxt)
            xmlDoc.selectSingleNode("/config/CHK/normal/cyt").Text = CStr(gudtSpecData.intCHKNormalCyt)
            xmlDoc.selectSingleNode("/config/PRESETGAN/normal/R").Text = CStr(gudtSpecData.intPRESETGANNormalR)
            xmlDoc.selectSingleNode("/config/PRESETGAN/normal/G").Text = CStr(gudtSpecData.intPRESETGANNormalG)
            xmlDoc.selectSingleNode("/config/PRESETGAN/normal/B").Text = CStr(gudtSpecData.intPRESETGANNormalB)
            xmlDoc.selectSingleNode("/config/PRESETOFF/normal/R").Text = CStr(gudtSpecData.intPRESETOFFNormalR)
            xmlDoc.selectSingleNode("/config/PRESETOFF/normal/G").Text = CStr(gudtSpecData.intPRESETOFFNormalG)
            xmlDoc.selectSingleNode("/config/PRESETOFF/normal/B").Text = CStr(gudtSpecData.intPRESETOFFNormalB)
        ElseIf strColorTemp = COLORTEMP_WARM1 Then
            xmlDoc.selectSingleNode("/config/SPEC/warm1/x").Text = CStr(gudtSpecData.intSPECWarm1x)
            xmlDoc.selectSingleNode("/config/SPEC/warm1/y").Text = CStr(gudtSpecData.intSPECWarm1y)
            xmlDoc.selectSingleNode("/config/TOL/warm1/xt").Text = CStr(gudtSpecData.intTOLWarm1xt)
            xmlDoc.selectSingleNode("/config/TOL/warm1/yt").Text = CStr(gudtSpecData.intTOLWarm1yt)
            xmlDoc.selectSingleNode("/config/CHK/warm1/cxt").Text = CStr(gudtSpecData.intCHKWarm1Cxt)
            xmlDoc.selectSingleNode("/config/CHK/warm1/cyt").Text = CStr(gudtSpecData.intCHKWarm1Cyt)
            xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/R").Text = CStr(gudtSpecData.intPRESETGANWarm1R)
            xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/G").Text = CStr(gudtSpecData.intPRESETGANWarm1G)
            xmlDoc.selectSingleNode("/config/PRESETGAN/warm1/B").Text = CStr(gudtSpecData.intPRESETGANWarm1B)
            xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/R").Text = CStr(gudtSpecData.intPRESETOFFWarm1R)
            xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/G").Text = CStr(gudtSpecData.intPRESETOFFWarm1G)
            xmlDoc.selectSingleNode("/config/PRESETOFF/warm1/B").Text = CStr(gudtSpecData.intPRESETOFFWarm1B)
        End If
    
        xmlDoc.selectSingleNode("/config/SPEC/cool1/Lv").Text = CStr(gudtSpecData.intSPECCool1Lv)
        xmlDoc.selectSingleNode("/config/SPEC/normal/Lv").Text = CStr(gudtSpecData.intSPECNormalLv)
        xmlDoc.selectSingleNode("/config/SPEC/warm1/Lv").Text = CStr(gudtSpecData.intSPECWarm1Lv)
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
