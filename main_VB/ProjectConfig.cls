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
    mstrConfigFilePath = App.Path & "\" & strCurrentModelName & "\config.xml"

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
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean

    If Not CBool(PathFileExists(mstrConfigFilePath)) Then
        MsgBox "Cannot open " & mstrConfigFilePath & " file."
        End
    End If
    
    success = xmlDoc.Load(mstrConfigFilePath)
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        If xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "UART" Then
            mConfigData.CommMode = modeUART
        Else
            mConfigData.CommMode = modeNetwork
        End If

        mConfigData.strComBaud = xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@baud").Text
        mConfigData.intComID = Val(xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@id").Text)
        mConfigData.strInputSource = xmlDoc.selectSingleNode("/config/input_source").Text
        mConfigData.lngDelayMs = Val(xmlDoc.selectSingleNode("/config/delayms").Text)
        mConfigData.intChannelNum = Val(xmlDoc.selectSingleNode("/config/channel_number").Text)
        mConfigData.intBarCodeLen = Val(xmlDoc.selectSingleNode("/config/length_bar_code").Text)
        mConfigData.intLvSpec = Val(xmlDoc.selectSingleNode("/config/Lv_spec").Text)
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
    
    success = xmlDoc.Load(mstrConfigFilePath)
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        If mConfigData.CommMode = modeUART Then
            xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "UART"
        Else
            xmlDoc.selectSingleNode("/config/communication").selectSingleNode("@mode").Text = "Network"
        End If
        
        xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@baud").Text = mConfigData.strComBaud
        xmlDoc.selectSingleNode("/config/communication/common").selectSingleNode("@id").Text = CStr(mConfigData.intComID)
        xmlDoc.selectSingleNode("/config/input_source").Text = mConfigData.strInputSource
        xmlDoc.selectSingleNode("/config/delayms").Text = CStr(mConfigData.lngDelayMs)
        xmlDoc.selectSingleNode("/config/channel_number").Text = CStr(mConfigData.intChannelNum)
        xmlDoc.selectSingleNode("/config/length_bar_code").Text = CStr(mConfigData.intBarCodeLen)
        xmlDoc.selectSingleNode("/config/Lv_spec").Text = CStr(mConfigData.intLvSpec)
        
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
        
        xmlDoc.Save mstrConfigFilePath
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

Public Property Get barCodeLen() As Integer
    barCodeLen = mConfigData.intBarCodeLen
End Property

Public Property Let barCodeLen(intBarCodeLen As Integer)
    mConfigData.intBarCodeLen = intBarCodeLen
End Property

Public Property Get LvSpec() As Integer
    LvSpec = mConfigData.intLvSpec
End Property

Public Property Let LvSpec(intLvSpec As Integer)
    mConfigData.intLvSpec = intLvSpec
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
