Public Class frm_tagStyle_Map
    Public targetDoc As Word.Document
    Public Sub New(ByRef myDoc As Word.Document)
        ' This call is required by the designer.
        InitializeComponent()
        '
        Me.targetDoc = myDoc
        '
        Me.frm_refresh_tagStyleMap(Me.targetDoc)
        '
        ' Add any initialization after the InitializeComponent() call.
    End Sub
    '
    Public Function frm_refresh_tagStyleMap(ByRef myDoc As Word.Document) As Collection
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim lstOfSections As Collection
        Dim strRslt As String

        '
        lstOfSections = objHfMgr.hf_getTagStyleMap_All(myDoc)
        Me.lstBx_docMap.Items.Clear()
        '
        If lstOfSections.Count > 0 Then
            For j = 1 To lstOfSections.Count
                strRslt = lstOfSections.Item(CStr(j))
                Me.lstBx_docMap.Items.Add(strRslt)
            Next
        Else

        End If
        '
        Return lstOfSections
        '
    End Function

    Private Sub btn_Close_Click(sender As Object, e As EventArgs) Handles btn_Close.Click
        Me.Close()
        '
    End Sub

    Private Sub btn_refresh_Click(sender As Object, e As EventArgs) Handles btn_refresh.Click
        Me.frm_refresh_tagStyleMap(Me.targetDoc)
    End Sub

    Private Sub ctx_menuItem_GoToSection_Click(sender As Object, e As EventArgs) Handles ctx_menuItem_GoToSection.Click
        Dim sect As Word.Section
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        Dim j As Integer
        '
        j = Me.lstBx_docMap.SelectedIndex
        '
        'Remember the list is zero based while sections start at 1
        '
        Try
            sect = Me.targetDoc.Sections.Item(j + 1)
            rng = sect.Range
            para = rng.Paragraphs.Item(1)
            rng = para.Range
            rng.MoveEnd(, -1)
            '
            If Not rng.Text = "" Then
                rng.Select()
            Else
                'If the first line of the section is empty, then select the
                'contents of the section
                rng = sect.Range
                rng.MoveEnd(, -2)
                rng.Select()
            End If
            '
        Catch ex As Exception

        End Try

    End Sub

End Class