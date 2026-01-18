Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cBrandMgr
    Public objGlobals As cGlobals
    Public Sub New()
        Me.objGlobals = New cGlobals()
    End Sub
    '
    ''' <summary>
    ''' This method will rebuild the background of sect, where sect is the section is a 'page'  
    ''' of some sort
    ''' </summary>
    ''' <param name="sect"></param>
    Public Overridable Sub brnd_Rebuild_Background(ByRef sect As Word.Section, deleteHeaderFooters As Boolean, doLogo As Boolean, Optional strTagStyle As String = "tag_coverPage")
        Dim objHFMgr As New cHeaderFooterMgr
        Dim hf As Word.HeaderFooter
        Dim rng As Word.Range
        Dim shp As Word.Shape
        Dim objBBMgr As New cBBlocksHandler()
        '
        If deleteHeaderFooters Then objHFMgr.hf_hfs_deleteAll(sect)
        '
        If sect.PageSetup.DifferentFirstPageHeaderFooter Then
            hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        Else
            hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        End If
        rng = hf.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        If deleteHeaderFooters Then
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
        Else
            If hf.Range.Tables.Count <> 0 Then
                rng = hf.Range.Tables.Item(hf.Range.Tables.Count).Range
                'rng = hf.Range.Tables.Item(1).Range.Cells.Item(1).Range
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            Else
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
            End If
        End If
        '
        'We need to establish the  tag Style.. in the Header
        '
        'rng.Style = objGlobals.glb_get_wrdActiveDoc.Styles(strTagStyle)
        objHFMgr.hf_tags_setTagStyle(sect, strTagStyle)

        '
        shp = objHFMgr.hf_Insert_BackShape(hf, rng)
        '
        If doLogo Then Me.brnd_Insert_Logo(hf)
        '
    End Sub
    '
#Region "Logo Handling"
    ''' <summary>
    ''' This method will insert the aac logo at the beginning of the HeaderFooter hf.
    ''' The logo is extracted from the template's building blocks library. The logo is
    ''' returned for post processing as required
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <returns></returns>
    Public Function brnd_Insert_Logo(ByRef hf As Word.HeaderFooter) As Word.Shape
        Dim strLogoName As String = "aac_Cpg_Logo"
        Dim objBBMgr As New cBBlocksHandler()
        Dim objWCAGMgr As New cWCAGMgr()
        Dim rng As Word.Range
        Dim shp As Word.Shape
        '
        shp = Nothing
        '
        'First delete any logo artefacts
        If hf.Shapes.Count <> 0 Then
            For Each shp In hf.Shapes
                If shp.Name = strLogoName Then
                    shp.Delete()
                    'Exit For
                End If
            Next
        End If
        rng = hf.Range
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)             'Just in case there are tables in the header
        '
        Try
            rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange(strLogoName, "CoverPage", rng)
            shp = rng.ShapeRange.Item(1)
            shp.Name = strLogoName
            shp.Left = 57.0
            shp.Top = 57.0
            '
            objWCAGMgr.wcag_set_decorative(shp, True)
            '
        Catch ex As Exception
            shp = Nothing
        End Try
        '
        Return shp
        '
    End Function
    '
    Public Sub brnd_recolour_Logo(ByRef hf As Word.HeaderFooter)
        Dim shp, grpItem As Word.Shape
        '
        Try
            shp = hf.Range.ShapeRange.Item("logo_AAC_TandG")
            For Each grpItem In shp.GroupItems
                If grpItem.Name Like "Freeform*" Then grpItem.Fill.ForeColor.RGB = RGB(255, 255, 255)
                If grpItem.Name Like "Rectangle*" Then grpItem.Fill.ForeColor.RGB = Me.objGlobals._glb_colour_purple_Mid
            Next
        Catch ex As Exception

        End Try
    End Sub
    '
    '
    Public Sub brnd_recolour_Logo(ByRef shp As Word.Shape)
        Dim grpItem As Word.Shape
        '
        Try
            For Each grpItem In shp.GroupItems
                If grpItem.Name Like "Freeform*" Then grpItem.Fill.ForeColor.RGB = RGB(255, 255, 255)
                If grpItem.Name Like "Rectangle*" Then grpItem.Fill.ForeColor.RGB = Me.objGlobals._glb_colour_purple_Mid
            Next
        Catch ex As Exception

        End Try
    End Sub
    '

    ''' <summary>
    ''' This method will insert the aac logo at the beginning of the range rng.
    ''' The logo is extracted from the template's building blocks library. The logo
    ''' is returned for post processing as required
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <returns></returns>
    Public Function brnd_Insert_Logo(ByRef rng As Word.Range) As Word.Shape
        Dim strLogoName As String = "aac_Cpg_Logo"
        Dim objBBMgr As New cBBlocksHandler()
        Dim shp As Word.Shape
        '
        shp = Nothing
        'First delete any logo artefacts
        If rng.ShapeRange.Count <> 0 Then
            For Each shp In rng.ShapeRange
                If shp.Name = strLogoName Then
                    shp.Delete()
                    'Exit For
                End If
            Next
        End If
        '
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)             'Just in case there are tables in the header
        '
        Try
            rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange(strLogoName, "CoverPage", rng)
            shp = rng.ShapeRange.Item(1)
            shp.Name = strLogoName
            shp.Left = 57.0
            shp.Top = 57.0
        Catch ex As Exception
            shp = Nothing
        End Try
        '
        Return shp
        '
    End Function
    '
#End Region

#Region "Landscape Chapter Banner"

    Public Sub brnd_LndScp_ChptBanner_jigsaw(ByRef sect As Word.Section)
        Dim hf As Word.HeaderFooter
        Dim shp As Word.Shape
        Dim rng As Word.Range
        Dim objBBMgr As New cBBlocksHandler()
        Dim strBuildingBlockName, strCategoryName As String
        '
        strBuildingBlockName = "aac_jigSaw_Wide"
        strCategoryName = "documentBody"
        '
        hf = sect.Headers.Item(1)
        rng = hf.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange(strBuildingBlockName, strCategoryName, rng)
        '
        'rng = MyBase.insert_PreBuilt_Elements(rng, "aac_jigSaw_Wide", "documentBody")
        If rng.ShapeRange.Count <> 0 Then
            shp = rng.ShapeRange.Item(1)
            shp.LockAspectRatio = False
            shp.Width = 718.5
            shp.Height = 241.1
            shp.LockAspectRatio = True
            shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
            shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
            shp.Top = 295.5
            shp.Left = 0.0
        End If
        '
    End Sub

#End Region

End Class
