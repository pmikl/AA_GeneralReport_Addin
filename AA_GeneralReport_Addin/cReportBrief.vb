Imports Microsoft.Office.Interop.Word
Public Class cReportBrief
    Inherits cChptBase
    '
    '
    Public Sub New()
        MyBase.New()
    End Sub
    '
    ''' <summary>
    ''' This method will determine whether the current document is a Brief.
    ''' It does so by checking for the 'tag_aaBrief' style in the header
    ''' of the first section
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function brf_is_brief(ByRef myDoc As Word.Document) As Boolean
        Dim rslt As Boolean
        Dim strTag As String
        Dim objHfMgr As New cHeaderFooterMgr()
        '
        rslt = False
        strTag = objHfMgr.hf_tags_getTagStyleName(myDoc.Sections.First)
        '
        If strTag = "tag_aaBrief" Then rslt = True
        '
        Return rslt
        '
    End Function

    '
    ''' <summary>
    ''' This method will test a section for the 'Brief' tag 'tag_aaBrief' and will
    ''' return true if it finds one in the header
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function brf_section_isFromBrief(ByRef sect As Word.Section) As Boolean
        Dim rslt As Boolean
        Dim strTag As String
        Dim objHfMgr As New cHeaderFooterMgr()
        '
        rslt = False
        strTag = objHfMgr.hf_tags_getTagStyleName(sect)
        '
        If strTag = "tag_aaBrief" Then rslt = True
        '
        Return rslt
        '
    End Function

    '
    ''' <summary>
    ''' This method will clear out all items in the first and primary page header/footers
    ''' that do not have the name "logo_AAC*"
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub brf_delete_backItems_And_FixLogos(ByRef sect As Word.Section)
        Dim hf As Word.HeaderFooter
        Dim objLogos As New cLogosMgr()
        '       
        Try
            If Me.brf_is_brief(objGlobals.glb_get_wrdActiveDoc) Then
                objLogos.logos_set_colour(sect, RGB(0, 0, 0), -1)

                If sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary).Exists Then
                    hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                    Me.brf_delete_backItems_exceptLogos(hf, "logo_AAC*")
                End If
                '
                If sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Exists Then
                    hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                    Me.brf_delete_backItems_exceptLogos(hf, "logo_AAC*")
                End If
            End If
            '
        Catch ex As Exception

        End Try
        '
    End Sub
    '
    ''' <summary>
    ''' This method will delete all shapes in the HeaderFooter (hf) that have a name that is not like
    ''' strShpName. The default is strShpName = 'logo_AAC*'.. This is typically used to clear out the
    ''' first page hf items that are found in a new section derived from the first page of the AA Brief
    ''' </summary>
    ''' <param name="hf"></param>
    Public Sub brf_delete_backItems_exceptLogos(ByRef hf As Word.HeaderFooter, Optional strShpName As String = "logo_AAC*")
        Dim rng As Word.Range
        Dim shp As Word.Shape
        Dim j As Integer
        '
        Try
            rng = hf.Range
            If rng.ShapeRange.Count <> 0 Then
                For j = rng.ShapeRange.Count To 1 Step -1
                    shp = rng.ShapeRange.Item(j)
                    If Not shp.Name Like strShpName Then
                        shp.Delete()
                    End If
                Next
            End If
            '
        Catch ex As Exception

        End Try
        '
    End Sub

End Class
