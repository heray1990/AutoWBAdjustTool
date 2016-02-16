VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProjectConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************
' Class module for handling config.xml of the
' application.
'**********************************************

Option Explicit

Public Enum CommunicationMode
    modeUART = 1
    modeNetwork
End Enum

Public Type udtConfigData
    CommMode As CommunicationMode
    strCommBaud As String
    intCommID As Integer
    strInputSource As String
    lngDelayMs As Long
    intChannelNum As Integer
    intBarCodeLen As Integer
    intLvSpec As Integer
    bolEnableCool2 As Boolean
    bolEnableCool1 As Boolean
    bolEnableNormal As Boolean
    bolEnableWarm1 As Boolean
    bolEnableWarm2 As Boolean
    bolEnableChkColor As Boolean
    bolEnableAdjOffset As Boolean
End Type

Private mConfigData As udtConfigData
Private mstrConfigFilePath As String

Private Sub Class_Initialize()
    mstrConfigFilePath = strCurrentModelName & "\config.xml"
    mConfigData.CommMode = modeUART
    mConfigData.strCommBaud = "115200"
    mConfigData.intCommID = 1
    mConfigData.strInputSource = "HDMI1"
    mConfigData.lngDelayMs = 500
    mConfigData.intChannelNum = 1
    mConfigData.intBarCodeLen = 1
    mConfigData.intLvSpec = 280
    mConfigData.bolEnableCool2 = False
    mConfigData.bolEnableCool1 = True
    mConfigData.bolEnableNormal = True
    mConfigData.bolEnableWarm1 = True
    mConfigData.bolEnableWarm2 = False
    mConfigData.bolEnableChkColor = True
    mConfigData.bolEnableAdjOffset = False
End Sub

Public Sub LoadConfigData()
    LoadCommMode

    mConfigData.strInputSource = GetXmlNodeValue(mstrConfigFilePath, "/config/input_source")
    mConfigData.lngDelayMs = Val(GetXmlNodeValue(mstrConfigFilePath, "/config/delayms"))
    mConfigData.intChannelNum = Val(GetXmlNodeValue(mstrConfigFilePath, "/config/channel_number"))
    mConfigData.intBarCodeLen = Val(GetXmlNodeValue(mstrConfigFilePath, "/config/length_bar_code"))
    mConfigData.intLvSpec = Val(GetXmlNodeValue(mstrConfigFilePath, "/config/Lv_spec"))

    If GetXmlNodeValue(mstrConfigFilePath, "/config/cool_2") = "True" Then
        mConfigData.bolEnableCool2 = True
    Else
        mConfigData.bolEnableCool2 = False
    End If
    
    If GetXmlNodeValue(mstrConfigFilePath, "/config/cool_1") = "True" Then
        mConfigData.bolEnableCool1 = True
    Else
        mConfigData.bolEnableCool1 = False
    End If
    
    If GetXmlNodeValue(mstrConfigFilePath, "/config/normal") = "True" Then
        mConfigData.bolEnableNormal = True
    Else
        mConfigData.bolEnableNormal = False
    End If
    
    If GetXmlNodeValue(mstrConfigFilePath, "/config/warm_1") = "True" Then
        mConfigData.bolEnableWarm1 = True
    Else
        mConfigData.bolEnableWarm1 = False
    End If
    
    If GetXmlNodeValue(mstrConfigFilePath, "/config/warm_2") = "True" Then
        mConfigData.bolEnableWarm2 = True
    Else
        mConfigData.bolEnableWarm2 = False
    End If
    
    If GetXmlNodeValue(mstrConfigFilePath, "/config/check_color") = "True" Then
        mConfigData.bolEnableChkColor = True
    Else
        mConfigData.bolEnableChkColor = False
    End If
    
    If GetXmlNodeValue(mstrConfigFilePath, "/config/adjust_offset") = "True" Then
        mConfigData.bolEnableAdjOffset = True
    Else
        mConfigData.bolEnableAdjOffset = False
    End If
End Sub

Private Sub LoadCommMode()
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(App.Path & "\" & strCurrentModelName & "\config.xml")
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        Dim objNode As MSXML2.IXMLDOMNode
        
        Set objNode = xmlDoc.selectSingleNode("/config/communication")
        
        If objNode Is Nothing Then
            MsgBox "There is not <communication> node in config.xml."
        Else
            If objNode.selectSingleNode("@mode").Text = "UART" Then
                mConfigData.CommMode = modeUART
                LoadCommSetting
            Else
                mConfigData.CommMode = modeNetwork
            End If
        End If
    End If
End Sub

Private Sub LoadCommSetting()
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(App.Path & "\" & strCurrentModelName & "\config.xml")
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        Dim objNode As MSXML2.IXMLDOMNode
        
        Set objNode = xmlDoc.selectSingleNode("/config/communication/common")
        
        If objNode Is Nothing Then
            MsgBox "There is not <common> node in config.xml."
        Else
            mConfigData.strCommBaud = objNode.selectSingleNode("@baud").Text
            mConfigData.intCommID = Val(objNode.selectSingleNode("@id").Text)
        End If
    End If
End Sub

