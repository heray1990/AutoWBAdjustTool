Option Strict Off
Option Explicit On
Module WizardXMLHelper
	'**********************************************
	' Handling wizard.xml for the application
	'**********************************************
	
	
	Public Const gstrDelimiterForProjName As String = "-"
	
	' Return current project's name.
	Public Function GetCurProjectName() As String
		Dim xmlDoc As New MSXML2.DOMDocument
		Dim success As Boolean
		
		success = xmlDoc.Load(My.Application.Info.DirectoryPath & "\wizard.xml")
		
		Dim objNode As MSXML2.IXMLDOMNode
		If success = False Then
			MsgBox(xmlDoc.parseError.reason)
		Else
			
			objNode = xmlDoc.selectSingleNode("/wizard/current_project")
			
			If objNode Is Nothing Then
				MsgBox("There is not <current_project> node in wizard.xml.")
				GetCurProjectName = "???"
			Else
				GetCurProjectName = objNode.Text
			End If
		End If
	End Function
	
	' Save current project's name.
	Public Sub SetCurProjectName(ByRef strCurProjectName As String)
		Dim xmlDoc As New MSXML2.DOMDocument
		Dim success As Boolean
		
		success = xmlDoc.Load(My.Application.Info.DirectoryPath & "\wizard.xml")
		
		Dim objNode As MSXML2.IXMLDOMNode
		If success = False Then
			MsgBox(xmlDoc.parseError.reason)
		Else
			
			objNode = xmlDoc.selectSingleNode("/wizard/current_project")
			objNode.Text = strCurProjectName
			
			xmlDoc.Save(My.Application.Info.DirectoryPath & "\wizard.xml")
		End If
	End Sub
	
	' Return the list of projects' name.
	Public Function GetProjectList() As Collection
		Dim xmlDoc As New MSXML2.DOMDocument
		Dim success As Boolean
		Dim colProjectList As Collection
		
		colProjectList = New Collection
		
		success = xmlDoc.Load(My.Application.Info.DirectoryPath & "\wizard.xml")
		
		Dim objNodeList As MSXML2.IXMLDOMNodeList
		Dim objNode As MSXML2.IXMLDOMNode
		Dim brand As Object
		Dim model As String
		If success = False Then
			MsgBox(xmlDoc.parseError.reason)
		Else
			
			objNodeList = xmlDoc.selectNodes("/wizard/project_list/project")
			
			If Not objNodeList Is Nothing Then
				
				For	Each objNode In objNodeList
					'UPGRADE_WARNING: Couldn't resolve default property of object brand. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					brand = Trim(objNode.selectSingleNode("@brand").Text)
					model = Trim(objNode.selectSingleNode("@model").Text)
					'UPGRADE_WARNING: Couldn't resolve default property of object brand. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					colProjectList.Add(brand & gstrDelimiterForProjName & model)
				Next objNode
			End If
		End If
		
		GetProjectList = colProjectList
		'UPGRADE_NOTE: Object colProjectList may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		colProjectList = Nothing
	End Function
End Module