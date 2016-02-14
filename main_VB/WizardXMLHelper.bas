Attribute VB_Name = "WizardXMLHelper"
'**********************************************
' Handling wizard.xml for the application
'**********************************************

Option Explicit

' Return current project's name.
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
    ByVal strNodeName As String, _
    Optional ByVal strDefaultValue As String = "") As String

    Dim ValueNode As MSXML2.IXMLDOMNode

    Set ValueNode = StartAtNode.selectSingleNode(".//" & strNodeName)
    If ValueNode Is Nothing Then
        GetNodeValue = strDefaultValue
    Else
        GetNodeValue = ValueNode.Text
    End If
End Function

' Save current project's name.
Public Sub SetCurProjectName(strCurProjectName As String)
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(App.Path & "\wizard.xml")
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        Dim objNode As MSXML2.IXMLDOMNode
        
        Set objNode = xmlDoc.selectSingleNode("/wizard/current_project")
        objNode.Text = strCurProjectName
        
        xmlDoc.Save App.Path & "\wizard.xml"
    End If
End Sub

' Return the list of projects' name.
Public Function GetProjectList() As Collection
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    Dim colProjectList As Collection
    
    Set colProjectList = New Collection
    
    success = xmlDoc.Load(App.Path & "\wizard.xml")
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        Dim objNodeList As MSXML2.IXMLDOMNodeList
        
        Set objNodeList = xmlDoc.selectNodes("/wizard/project_list/project")
        
        If Not objNodeList Is Nothing Then
            Dim objNode As MSXML2.IXMLDOMNode
            Dim brand, model As String
            
            For Each objNode In objNodeList
                brand = objNode.selectSingleNode("@brand").Text
                model = objNode.selectSingleNode("@model").Text
                colProjectList.Add brand & "-" & model
            Next objNode
        End If
    End If
    
    Set GetProjectList = colProjectList
    Set colProjectList = Nothing
End Function
