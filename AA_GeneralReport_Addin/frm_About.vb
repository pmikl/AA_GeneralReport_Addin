Public Class frm_About
    Dim objGlobals As New cGlobals()
    Public Sub New()
        Dim lstOfVersionNumbers As Collection
        Dim strPublishVersion, strTmplVersion As String
        Dim prop As Microsoft.Office.Core.DocumentProperty
        Dim tmpl As Word.Template

        ' This call is required by the designer.
        InitializeComponent()
        strTmplVersion = "none"
        '
        tmpl = objGlobals.glb_get_wrdActiveDoc.AttachedTemplate
        '
        ' Add any initialization after the InitializeComponent() call.
        Try
            lstOfVersionNumbers = objGlobals.glb_get_VersionInformation()
            strPublishVersion = CStr(lstOfVersionNumbers("major")) + "." + CStr(lstOfVersionNumbers("minor")) + "."
            strPublishVersion = strPublishVersion + CStr(lstOfVersionNumbers("build")) + "." + CStr(lstOfVersionNumbers("revision"))
            '
            Me.lbl_PublishVersion.Text = "Addin Version: " + strPublishVersion
            Me.txtBox_upDateSite.Text = CStr(lstOfVersionNumbers("updateSite"))
            '
            Try
                prop = tmpl.BuiltInDocumentProperties("Category")
                strTmplVersion = CStr(prop.Value)
                tmpl.Saved = True
            Catch ex2 As Exception
                strTmplVersion = "unknown"
            End Try
            '
            Me.lbl_TemplateVersion.Text = "Template version: " + strTmplVersion
            Me.txtBox_attachedTemplate.Text = tmpl.FullName
            '
            'Position the label
            Me.lbl_firstReleaseDate.Left = Me.btn_OK.Left + Me.btn_OK.Width - Me.lbl_firstReleaseDate.Width
            '
        Catch ex As Exception

        End Try
        '
    End Sub
    '
    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click
        Me.Close()
    End Sub
End Class