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

Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Private Type udtConfigData
    CommMode As CommunicationMode
    strComBaud As String
    intComID As Integer
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
    mConfigData.CommMode = modeUART
    mConfigData.strComBaud = "115200"
    mConfigData.intComID = 1
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
    mstrConfigFilePath = App.Path & "\" & strCurrentModelName & "\config.xml"
    
    If Not CBool(PathFileExists(mstrConfigFilePath)) Then
        MsgBox "Cannot open " & mstrConfigFilePath & " file."
        End
    End If
    
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
    
    success = xmlDoc.Load(mstrConfigFilePath)
    
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
    
    success = xmlDoc.Load(mstrConfigFilePath)
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        Dim objNode As MSXML2.IXMLDOMNode
        
        Set objNode = xmlDoc.selectSingleNode("/config/communication/common")
        
        If objNode Is Nothing Then
            MsgBox "There is not <common> node in config.xml."
        Else
            mConfigData.strComBaud = objNode.selectSingleNode("@baud").Text
            mConfigData.intComID = Val(objNode.selectSingleNode("@id").Text)
        End If
    End If
End Sub

Public Property Get CommMode() As CommunicationMode
   CommMode = mConfigData.CommMode
End Property

Public Property Get ComBaud() As String
   ComBaud = mConfigData.strComBaud
End Property

Public Property Get ComID() As Integer
   ComID = mConfigData.intComID
End Property

Public Property Get InputSource() As String
   InputSource = mConfigData.strInputSource
End Property

Public Property Get DelayMs() As Long
   DelayMs = mConfigData.lngDelayMs
End Property

Public Property Get ChannelNum() As Integer
   ChannelNum = mConfigData.intChannelNum
End Property

Public Property Get BarCodeLen() As Integer
   BarCodeLen = mConfigData.intBarCodeLen
End Property

Public Property Get LvSpec() As Integer
   LvSpec = mConfigData.intLvSpec
End Property

Public Property Get EnableCool2() As Boolean
   EnableCool2 = mConfigData.bolEnableCool2
End Property

Public Property Get EnableCool1() As Boolean
   EnableCool1 = mConfigData.bolEnableCool1
End Property

Public Property Get EnableNormal() As Boolean
   EnableNormal = mConfigData.bolEnableNormal
End Property

Public Property Get EnableWarm1() As Boolean
   EnableWarm1 = mConfigData.bolEnableWarm1
End Property

Public Property Get EnableWarm2() As Boolean
   EnableWarm2 = mConfigData.bolEnableWarm2
End Property

Public Property Get EnableChkColor() As Boolean
   EnableChkColor = mConfigData.bolEnableChkColor
End Property

Public Property Get EnableAdjOffset() As Boolean
   EnableAdjOffset = mConfigData.bolEnableAdjOffset
End Property
