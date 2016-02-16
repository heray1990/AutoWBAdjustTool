Attribute VB_Name = "XMLHelper"
'**********************************************
' Handling xml files for the application
'**********************************************

Option Explicit

Public Function GetXmlNodeValue(strXmlFilePath As String, strNodeName As String) As String
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(App.Path & "\" & strXmlFilePath)
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        Dim objNode As MSXML2.IXMLDOMNode
        
        Set objNode = xmlDoc.selectSingleNode(strNodeName)
        
        If objNode Is Nothing Then
            MsgBox "There is not " & strNodeName & " node in " & strXmlFilePath & " file."
        Else
            GetXmlNodeValue = objNode.Text
        End If
    End If
End Function

