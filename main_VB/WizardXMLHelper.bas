Attribute VB_Name = "WizardXMLHelper"
'**********************************************
' Handling wizard.xml for the application
'**********************************************

Option Explicit

' Return the name of current project.
Public Function GetCurProjectName() As String
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(App.Path & "\wizard.xml")
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        Dim objNode As MSXML2.IXMLDOMNode
        
        Set objNode = xmlDoc.selectSingleNode("wizard")
        GetCurProjectName = GetNodeValue(objNode, "current_project", "???")
    End If
End Function

' Return the node's value.
Private Function GetNodeValue(ByVal StartAtNode As IXMLDOMNode, _
    ByVal NodeName As String, _
    Optional ByVal DefaultValue As String = "") As String

    Dim ValueNode As MSXML2.IXMLDOMNode

    Set ValueNode = StartAtNode.selectSingleNode(".//" & NodeName)
    If ValueNode Is Nothing Then
        GetNodeValue = DefaultValue
    Else
        GetNodeValue = ValueNode.Text
    End If
End Function

