Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports System.IO
Imports System.IO.Packaging                     'Comes from the reference 'WindowsBase'
Imports System.Xml
Public Class cPropertyMgr
    'This software is the property of Peter Mikelaitis (http://mikl.com.au).
    'Peter Mikelaitis grants Acil Allen
    'a non exclusive and unrestricted licence to use the software
    '
    Public objGlobals As cGlobals
    Public strTrue As String
    Public strFalse As String
    Public strNullValue As String
    Public strWCAG_font_std As String
    Public strWCAG_font_11pt As String
    '
    '
    Public strWarningsOffProperty As String
    '
    Public Sub New()
        Me.objGlobals = New cGlobals()
        Me.strTrue = "true"
        Me.strFalse = "false"
        Me.strNullValue = "nothing"
        '
        Me.strWarningsOffProperty = "warningsoff"
        '
        Me.strWCAG_font_11pt = "11pt"
        Me.strWCAG_font_std = "std"
        '
        Me.objGlobals = New cGlobals()
        '
    End Sub
    '
    ''' <summary>
    ''' This method will return true if a specific CustomPrperty exists.. Typically used in the Doc.Open event
    ''' to test whether the document has ribbon related properties
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="strPropertyName"></param>
    ''' <returns></returns>
    Public Function prps_CustomProperty_Exists(ByRef myDoc As Word.Document, strPropertyName As String) As Boolean
        Dim rslt As Boolean = True
        Dim Props As Microsoft.Office.Core.DocumentProperties
        Dim prop As DocumentProperty
        Dim strTestValue As String = ""
        '
        '
        Try
            Props = myDoc.CustomDocumentProperties
            prop = Props.Item(strPropertyName)
            strTestValue = CStr(prop.Value)
            '
            rslt = True
            '
        Catch ex As Exception
            'We are here because there is no property. Since the Long report is the default we'll set
            'the property to Long report
            rslt = False
        End Try
        '
        Return rslt
    End Function


    '
    ''' <summary>
    ''' This method will set the font size to be used for 'Normal' in a document converted to
    ''' Accessible... Note that typically a number of other styles will be varied in wcag_styles_setForWCAG
    ''' to keep everything in proprtion. Values for strFontSize are; 'std' and '11pt'
    ''' 
    ''' </summary>
    ''' <param name="strFontSize"></param>
    ''' <returns></returns>
    Public Function props_wcag_setFontSize(Optional strFontSize As String = "std") As Boolean
        Dim rslt As Boolean = False
        '
        Select Case strFontSize
            Case Me.strWCAG_font_std, Me.strWCAG_font_11pt
            Case Else
                strFontSize = Me.strWCAG_font_std
        End Select
        '
        Try
            Me.prps_CustomProperty_set(strFontSize, "wcagFontSize")
            rslt = True
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' This method will get the font size to be used by the Accesibility Converter. If it can't find a setting it will
    ''' set an entry to Me.strWCAG_font_std. Vaid results are; 'std' and '11pt'
    ''' </summary>
    ''' <returns></returns>
    Public Function props_wcag_getFontSize() As String
        Dim rslt As String = ""
        '
        Try
            rslt = Me.prps_CustomProperty_get("wcagFontSize", Me.strWCAG_font_std)
        Catch ex As Exception
            rslt = Me.strWCAG_font_std
        End Try
        '
        Return rslt
        '
    End Function

    '
    '
    ''' <summary>
    ''' This method will return the Document Property value of the property
    ''' with the name strPropertyName. If the property does not exist, it will
    ''' create the property and set its valeu to strNullValue
    ''' </summary>
    ''' <param name="strPropertyName"></param>
    ''' <returns></returns>
    Public Function prps_CustomProperty_get(strPropertyName As String) As String
        '
        Dim Props As Microsoft.Office.Core.DocumentProperties
        Dim prop As DocumentProperty
        Dim myDoc As Word.Document
        '
        myDoc = Me.objGlobals.glb_get_wrdActiveDoc()
        '
        Try
            prps_CustomProperty_get = Me.strNullValue
            '
            Props = myDoc.CustomDocumentProperties
            prop = Props.Item(strPropertyName)
            prps_CustomProperty_get = CStr(prop.Value)
            '
        Catch ex As Exception
            'We are here because there is no property. Since the Long report is the default we'll set
            'the property to Long report
            Props = myDoc.CustomDocumentProperties
            Props.Add(strPropertyName, False, MsoDocProperties.msoPropertyTypeString, Me.strNullValue)
            prps_CustomProperty_get = Me.strNullValue
        End Try
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will return the Document Property value of the property
    ''' with the name strPropertyName. If the property does not exist, it will
    ''' create the property and set its value to strDefault
    ''' </summary>
    ''' <param name="strPropertyName"></param>
    ''' <returns></returns>
    Public Function prps_CustomProperty_get(strPropertyName As String, strDefault As String) As String
        '
        Dim Props As Microsoft.Office.Core.DocumentProperties
        Dim prop As DocumentProperty
        Dim myDoc As Word.Document
        '
        myDoc = Me.objGlobals.glb_get_wrdActiveDoc()
        '
        Try
            prps_CustomProperty_get = Me.strNullValue
            '
            Props = myDoc.CustomDocumentProperties
            prop = Props.Item(strPropertyName)
            prps_CustomProperty_get = CStr(prop.Value)
            '
        Catch ex As Exception
            'We are here because there is no property. Since the Long report is the default we'll set
            'the property to Long report
            Props = myDoc.CustomDocumentProperties
            Props.Add(strPropertyName, False, MsoDocProperties.msoPropertyTypeString, strDefault)
            prps_CustomProperty_get = strDefault
        End Try
        '
    End Function

    ''' <summary>
    ''' This method will return true if the specified property contains
    ''' the strNullValue string
    ''' </summary>
    ''' <param name="strPropertyName"></param>
    ''' <returns></returns>
    Public Function prps_prop_hasNoValue(strPropertyName As String) As Boolean
        Dim strPropertyValue As String
        '
        prps_prop_hasNoValue = False
        strPropertyValue = Me.prps_CustomProperty_get(strPropertyName)
        If strPropertyValue = Me.strNullValue Then prps_prop_hasNoValue = True
        '
    End Function
    '
    ''' <summary>
    ''' This method will set the value of the Document Property with the name strPropertyName
    ''' to the string value specified by strPropertyValue.. If the property does not exist it
    ''' will create the property
    ''' </summary>
    ''' <param name="strPropertyValue"></param>
    ''' <param name="strPropertyName"></param>
    Public Sub prps_CustomProperty_set(strPropertyValue As String, strPropertyName As String)
        Dim Props As DocumentProperties
        Dim prop As DocumentProperty
        Dim myDoc As Word.Document

        '
        On Error GoTo finis
        myDoc = Me.objGlobals.glb_get_wrdActiveDoc
        '
        'Don't write the property if the ActiveDocument is a Template
        If objGlobals.glb_doc_isTemplate() Then Exit Sub
        '
        'Set the stored mode indicator to short
        Props = myDoc.CustomDocumentProperties
        'Props = Globals.ThisDocument.CustomDocumentProperties
        prop = Props.Item(strPropertyName)
        prop.Value = strPropertyValue
        '
        Exit Sub
finis:
        Call Props.Add(strPropertyName, False, MsoDocProperties.msoPropertyTypeString, strPropertyValue)
    End Sub
    '
    ''' <summary>
    ''' This method will set the value of the Document (myDoc) Property with the name strPropertyName
    ''' to the string value specified by strPropertyValue.. If the property does not exist it
    ''' will attempt to create the property
    ''' </summary>
    ''' <param name="strPropertyValue"></param>
    ''' <param name="strPropertyName"></param>
    ''' <param name="myDoc"></param>
    Public Sub prps_CustomProperty_set(strPropertyValue As String, strPropertyName As String, ByRef myDoc As Word.Document)
        Dim Props As DocumentProperties
        Dim prop As DocumentProperty
        '
        Props = myDoc.CustomDocumentProperties
        '
        Try
            prop = Props.Item(strPropertyName)
            prop.Value = strPropertyValue
            '
        Catch ex As Exception
            Try
                Call Props.Add(strPropertyName, False, MsoDocProperties.msoPropertyTypeString, strPropertyValue)
            Catch ex2 As Exception

            End Try
        End Try
        '
    End Sub
    '
    Public Sub prps_del_customProperty(strPropertyName As String, ByRef myDoc As Word.Document)
        Dim Props As DocumentProperties
        Dim prop As DocumentProperty
        '
        'Don't act if the ActiveDocument is a Template
        If objGlobals.glb_doc_isTemplate() Then Exit Sub
        '
        Try
            Props = myDoc.CustomDocumentProperties
            prop = Props.Item(strPropertyName)
            prop.Delete()
            '
        Catch ex As Exception

        End Try

    End Sub
    '
    '
    ''' <summary>
    ''' This method will delete the Assembly references to the Ribbon... Tested OK 20220404
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub prps_rbn_del(ByRef myDoc As Word.Document)
        '
        Me.prps_del_customProperty("_AssemblyLocation", myDoc)
        Me.prps_del_customProperty("_AssemblyName", myDoc)
        '
        Try
            myDoc.AttachedTemplate = ""
        Catch ex As Exception

        End Try
        '
    End Sub
    '
    ''' <summary>
    ''' This method will return the value of the custom property '_AssemblyLocation'
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function prps_rbn_getAssemblyLocation(ByRef myDoc As Word.Document) As String
        Dim strAssemblyLocation As String
        '
        strAssemblyLocation = Me.prps_CustomProperty_get("_AssemblyLocation")
        '
        Return strAssemblyLocation
    End Function
    '
    ''' <summary>
    ''' This method will return the value of the custom property '_AssemblyName'
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function prps_rbn_getAssemblyName(ByRef myDoc As Word.Document) As String
        Dim strAssemblyName As String
        '
        strAssemblyName = Me.prps_CustomProperty_get("_AssemblyName")
        '
        Return strAssemblyName
    End Function

    '
    ''' <summary>
    ''' This method will put back (into myDoc) the Assembly references for the aac ribbon. We can select the
    ''' actual values depending on strClientId. If strClientId = "aac", then the references are set for the
    ''' Acil Allen deployment web site. If strClientId = "testMachine", then the references are set for the
    ''' test platform that pulls software off Shadow
    ''' </summary>
    ''' <param name="strClientId"></param>
    Public Sub prps_rbn_setReferences(strClientId As String, ByRef myDoc As Word.Document)
        'Dim tmpl As Word.Template
        '
        Select Case strClientId
            Case "aac"
                Me.prps_CustomProperty_set("http://templates.acilallen.com.au/word/report/install/AAC Report Template.vsto|c7720a94-7995-49d9-ae2f-a9a6c57d0dec", "_AssemblyLocation", myDoc)
                Me.prps_CustomProperty_set("4E3C66D5-58D4-491E-A7D4-64AF99AF6E8B", "_AssemblyName", myDoc)

            Case "mikl.net.au"
                Me.prps_CustomProperty_set("https://mikl.net.au/org_aa/office/word/GeneralReport/install/AA GeneralReport.vsto|38218949-3a6a-4064-898e-34119e91a4dc", "_AssemblyLocation", myDoc)
                Me.prps_CustomProperty_set("4E3C66D5-58D4-491E-A7D4-64AF99AF6E8B", "_AssemblyName", myDoc)
                '

            Case "testMachine"
                Me.prps_CustomProperty_set("file://shadow/Software/AAC_Templates/word/report/install/AAC Report Template.vsto|c7720a94-7995-49d9-ae2f-a9a6c57d0dec", "_AssemblyLocation", myDoc)
                Me.prps_CustomProperty_set("4E3C66D5-58D4-491E-A7D4-64AF99AF6E8B", "_AssemblyName", myDoc)
                '
            Case "generalReport"
                Me.prps_CustomProperty_set("http://templates.acilallen.com.au/word/GeneralReport/install/AAC GeneralReport.vsto|38218949-3a6a-4064-898e-34119e91a4dc", "_AssemblyLocation", myDoc)
                Me.prps_CustomProperty_set("4E3C66D5-58D4-491E-A7D4-64AF99AF6E8B", "_AssemblyName", myDoc)
                '
                Try
                    myDoc.AttachedTemplate = "C:\Templates\AA GeneralReport.dotx"
                Catch ex As Exception

                End Try
                '
            Case "generalReport_Internal"
                Me.prps_CustomProperty_set("file://shadow/Software/AAC_Templates/word/GeneralReport/install/AAC GeneralReport.vsto|38218949-3a6a-4064-898e-34119e91a4dc", "_AssemblyLocation", myDoc)
                Me.prps_CustomProperty_set("4E3C66D5-58D4-491E-A7D4-64AF99AF6E8B", "_AssemblyName", myDoc)

        End Select
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method manipulates the document xml to remaove all Custom document properties
    ''' </summary>
    ''' <param name="docPath"></param>
    Sub prps_Remove_CustomProperties_All(docPath As String)
        Dim package As System.IO.Packaging.Package
        Dim customPropsUri As Uri = New Uri("/docProps/custom.xml", UriKind.Relative)
        '
        package = Package.Open(docPath, FileMode.Open, FileAccess.ReadWrite)

        If package.PartExists(customPropsUri) Then
            package.DeletePart(customPropsUri)
        End If

        package.Close()
    End Sub
    '
    ''' <summary>
    ''' This method will remove the Custom Properties '_AssemblyName' and '_AssemblyLocation'
    ''' </summary>
    ''' <param name="docPath"></param>
    Sub prps_Remove_CustomProperties_AssemblyNameAndLocation(docPath As String)
        Dim customPropsUri As New Uri("/docProps/custom.xml", UriKind.Relative)

        Using package As Package = Package.Open(docPath, FileMode.Open, FileAccess.ReadWrite)
            If Not package.PartExists(customPropsUri) Then Exit Sub

            Dim part As PackagePart = package.GetPart(customPropsUri)
            Dim xmlDoc As New XmlDocument()
            Using stream As Stream = part.GetStream()
                xmlDoc.Load(stream)
            End Using

            Dim nsmgr As New XmlNamespaceManager(xmlDoc.NameTable)
            nsmgr.AddNamespace("cp", "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties")
            nsmgr.AddNamespace("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")

            Dim propsToRemove = {"_AssemblyName", "_AssemblyLocation"}
            Dim nodesToRemove As New List(Of System.Xml.XmlNode)

            For Each propName In propsToRemove
                Dim xpath = $"//cp:property[cp:name='{propName}']"
                Dim node = xmlDoc.SelectSingleNode(xpath, nsmgr)
                If node IsNot Nothing Then nodesToRemove.Add(node)
            Next

            For Each node In nodesToRemove
                node.ParentNode.RemoveChild(node)
            Next

            ' Overwrite the part with updated XML
            Using stream As Stream = part.GetStream(FileMode.Create, FileAccess.Write)
                xmlDoc.Save(stream)
            End Using
        End Using
    End Sub
    '
    Sub prps_Remove_RemoveCustomPropsFromOneDriveDoc()
        Dim docPath As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\OneDrive\myFiles\YourDocument.docx"
        Dim customPropsUri As New Uri("/docProps/custom.xml", UriKind.Relative)

        Using package As Package = Package.Open(docPath, FileMode.Open, FileAccess.ReadWrite)
            If Not package.PartExists(customPropsUri) Then Exit Sub

            Dim part As PackagePart = package.GetPart(customPropsUri)
            Dim xmlDoc As New XmlDocument()
            Using stream As Stream = part.GetStream()
                xmlDoc.Load(stream)
            End Using

            Dim nsmgr As New XmlNamespaceManager(xmlDoc.NameTable)
            nsmgr.AddNamespace("cp", "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties")
            nsmgr.AddNamespace("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")

            Dim propsToRemove = {"_AssemblyName", "_AssemblyLocation"}
            Dim nodesToRemove As New List(Of System.Xml.XmlNode)

            For Each propName In propsToRemove
                Dim xpath = $"//cp:property[cp:name='{propName}']"
                Dim node = xmlDoc.SelectSingleNode(xpath, nsmgr)
                If node IsNot Nothing Then nodesToRemove.Add(node)
            Next

            For Each node In nodesToRemove
                node.ParentNode.RemoveChild(node)
            Next

            Using stream As Stream = part.GetStream(FileMode.Create, FileAccess.Write)
                xmlDoc.Save(stream)
            End Using
        End Using
    End Sub

End Class
