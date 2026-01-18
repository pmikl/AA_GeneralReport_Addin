Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
'
''' <summary>
''' 
'''This class deals with all things related to the Chapter
'''Divider
'''
'''Peter Mikelaitis October 2020...http://mikl.com.au
'''
''' </summary>
Public Class cChptDivider
    Inherits cChptBase
    '
    Public tbl_DivBanner As Word.Table
    Public strTagStyleNameAP As String
    '
    Public Sub New()
        MyBase.New()
        '
        Dim objBanner As New cChptBanner()
        '
        Me.strTagStyleName = objBanner.bnr_get_tagStyles(objBanner.tag_div)
        Me.strTagStyleNameAP = objBanner.bnr_get_tagStyles(objBanner.tag_divAP)

        Me.objGlobals = New cGlobals()

    End Sub
    '
    ''' <summary>
    ''' This method will return true if the Divider type is a standard Chapter/Part divider
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function is_divider_Chpt(ByRef sect As Word.Section) As Boolean
        Dim rslt As Boolean
        Dim strDivType As String
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objBnrMgr As New cChptBanner()
        Dim strTagStyle
        '
        rslt = False
        strDivType = "error"
        '
        Try
            strTagStyle = objHfMgr.hf_tags_getTagStyleName(sect)
            If strTagStyle = objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_div) Then rslt = True
            '
            'rslt = Me.is_divider_Any(sect, strDivType)
            'If strDivType = "chpt" Then rslt = True
        Catch ex As Exception
            rslt = False
            strDivType = "error"

        End Try

        Return rslt
    End Function
    '    '
    ''' <summary>
    ''' This method will return true if the Divider type is a standard Appendix divider
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function is_divider_App(ByRef sect As Word.Section) As Boolean
        Dim rslt As Boolean
        Dim strDivType As String
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objBnrMgr As New cChptBanner()
        Dim strTagStyle
        '
        rslt = False
        strDivType = "error"
        '
        Try
            strTagStyle = objHfMgr.hf_tags_getTagStyleName(sect)
            If strTagStyle = objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_divAP) Then rslt = True

            '
            'rslt = Me.is_divider_Any(sect, strDivType)
            'If strDivType = "app" Then rslt = True

        Catch ex As Exception
            rslt = False
            strDivType = "error"
        End Try

        Return rslt
    End Function
    '
    '
    ''' <summary>
    ''' This method will return true if the sect is anyone of the valid Divider types. if it returns
    ''' true, the strDivTpe = 'app' or 'chpt' 
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function is_divider_Any(ByRef sect As Word.Section) As Boolean
        Dim rslt As Boolean
        Dim objHfMgr As New cHeaderFooterMgr()
        '
        rslt = False
        If Me.is_divider_App(sect) Or Me.is_divider_Chpt(sect) Then rslt = True
        '
        Return rslt
    End Function
    '
    ''' <summary>
    ''' This method will convert and existing page to a divider. The divider type is specified in strDividerType. It
    ''' can take on the values of cChptBanner.tag_div (objBnrMgr.tag_div) or cChptBanner.tag_divAP (objBnrMgr.tag_divAP)
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="strDividerType"></param>
    ''' <param name="strHeadingText"></param>
    Public Sub div_convert_toDivider(ByRef sect As Word.Section, strDividerType As String, Optional strHeadingText As String = "", Optional showFooter As Boolean = False)
        Dim lstOfDimensions As Collection
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objBnrMgr As New cChptBanner()
        Dim lstOfBannerSettings As Collection
        'Dim objBrndMgr As New cBrandMgr()
        Dim hf As Word.HeaderFooter
        Dim rng As Word.Range
        '
        lstOfDimensions = Me.objGlobals.glb_getDimensions_Divider_Prt()
        If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then
            lstOfDimensions = Me.objGlobals.glb_getDimensions_Divider_Lnd()
        End If
        '
        objGlobals.glb_setDimensions(sect, lstOfDimensions)
        '
        'Put in a standard chapter with Div banner settings then modify. Making certain that
        'the header is flush with the left margin
        lstOfBannerSettings = Nothing
        lstOfBannerSettings = objBnrMgr.bnr_get_BannerSettings(strDividerType, False)
        '
        'Override the default heading
        If strHeadingText <> "" Then
            lstOfBannerSettings.Remove("strHeadingtext")
            lstOfBannerSettings.Add(strHeadingText, "strHeadingtext")
        End If
        '
        objBnrMgr.bnr_insert_BannerBase(sect.Range, False, "", lstOfBannerSettings)
        '
        objHfMgr.hf_headers_insert(sect, -1, True, True)
        '
        If showFooter Then
            objHfMgr.hf_footers_insert(sect)
        Else
            objHfMgr.hf_footers_delete(sect)
        End If
        '
        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        rng = hf.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        objHfMgr.hf_Insert_BackShape(hf, rng)
        '
        objHfMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(strDividerType))
        '
        objHfMgr.hf_set_textColourFooter(sect, RGB(255, 255, 255))
        '

    End Sub


    '
    Public Sub div_convert_toAPDivider(ByRef sect As Word.Section)
        Dim lstOfDimensions As Collection
        Dim objHfMgr As New cHeaderFooterMgr()
        'Dim objBrndMgr As New cBrandMgr()
        Dim hf As Word.HeaderFooter
        Dim rng As Word.Range
        'Dim tbl As Word.Table
        '
        '
        lstOfDimensions = Me.objGlobals.glb_getDimensions_Divider_Prt()
        If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then
            lstOfDimensions = Me.objGlobals.glb_getDimensions_Divider_Lnd()
        End If
        '
        objGlobals.glb_setDimensions(sect, lstOfDimensions)
        '
        objHfMgr.hf_headers_insert(sect, -1, True, True)
        objHfMgr.hf_footers_delete(sect)
        '
        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        rng = hf.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        objHfMgr.hf_Insert_BackShape(hf, rng)
        '
        objHfMgr.hf_tags_setTagStyle(sect, Me.strTagStyleNameAP)
        '
        'tbl = Me.chptBase_insert_BannerBase()

        '

    End Sub
    '
    ''' <summary>
    ''' This method will insert a new "Report Body" divider. It will use the default Heading text
    ''' (located in cChptBanner.chptBnr_get_BannerSettings), or if strHeadingText is not null it will
    ''' use the value of strHeading as the Heading text. If "placebehind" is false it will place the Divider
    ''' in front of the current section. If it is true, it will place it in front of the current section
    ''' </summary>
    ''' <param name="placeBehind"></param>
    ''' <param name="strHeadingText"></param>
    ''' <returns></returns>
    Public Function div_insert_newBody(placeBehind As Boolean, ByRef sect As Word.Section, Optional strHeadingText As String = "", Optional showFooter As Boolean = False) As Word.Section
        Dim objRptMgr As New cReport()
        Dim objBnrMgr As New cChptBanner()
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objBrndMgr As New cBrandMgr()
        Dim lst As Collection
        Dim strRptMode, strDivType As String
        'Dim tbl As Word.Table
        'Dim sect As Word.Section
        Dim hf As Word.HeaderFooter
        Dim doImage As Boolean
        '
        doImage = False
        '
        strRptMode = objRptMgr.Rpt_Mode_Get()
        '
        strDivType = "prtDiv"
        If strRptMode = objRptMgr.rpt_isLnd Then strDivType = "lndDiv"
        '
        'Put in a standard chapter with Div banner settings then modify. Making certain that
        'the header is flush with the left margin
        lst = objBnrMgr.bnr_get_BannerSettings(objBnrMgr.tag_div, doImage)
        If strHeadingText <> "" Then
            lst.Remove("strHeadingtext")
            lst.Add(strHeadingText, "strHeadingtext")
        End If
        '
        'MyBase.chpt_Insert_Std(placeBehind, sect, lst, strRptMode, "prtDiv", 6)
        MyBase.chpt_Insert_Std(placeBehind, sect, lst, strRptMode, strDivType, 6, False)
        '
        If Not showFooter Then
            objHFMgr.hf_footers_delete(sect)
        End If
        '
        objHFMgr.hf_headers_insert(sect, -1.0)
        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        '
        '*** This one puts in Part xx
        'Me.div_edit_Table(Me.tbl_Banner)
        '
        objBrndMgr.brnd_recolour_Logo(hf)
        objBrndMgr.brnd_Rebuild_Background(sect, False, False, Me.strTagStyleName)
        '
        Me.chptBase_PageNumbering_Set(sect, False, 1, "div")
        '
        objHFMgr.hf_set_textColourFooter(sect, RGB(255, 255, 255))
        'objHFMgr.hf_set_textColourFooter(sect, RGB(255, 0, 0), "resetColour")
        '
        Return sect
        '
    End Function

    '
    '
    ''' <summary>
    ''' This method will insert a new "Appendix/Attachment" divider. It will use the default Heading text
    ''' (located in cChptBanner.chptBnr_get_BannerSettings), or if strHeadingText is not null it will 
    ''' use the value of strHeading as the Heading text. If "placebehind" is false it will place the Divider
    ''' in front of the current section. If it is true, it will place it in front of the current section.
    ''' On return the variable sect will contain the new section
    ''' </summary>
    ''' <param name="placeBehind"></param>
    ''' <param name="strHeadingText"></param>
    ''' <returns></returns>
    Public Function div_insert_newAP(placeBehind As Boolean, ByRef sect As Word.Section, strRptMode As String, Optional strHeadingText As String = "", Optional showFooter As Boolean = False) As Word.Section
        Dim objRptMgr As New cReport()
        Dim objBnrMgr As New cChptBanner()
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objBrndMgr As New cBrandMgr()
        Dim lst As Collection
        Dim tbl As Word.Table
        'Dim sect As Word.Section
        Dim hf As Word.HeaderFooter
        Dim doImage As Boolean
        Dim strDivType As String
        '
        doImage = False
        '
        strDivType = "prtDiv"
        If strRptMode = objRptMgr.rpt_isLnd Then strDivType = "lndDiv"
        '
        'Put in a standard chapter with Div banner settings then modify. Making certain that
        'the header is flush with the left margin
        lst = objBnrMgr.bnr_get_BannerSettings(objBnrMgr.tag_divAP, doImage)
        If strHeadingText <> "" Then
            lst.Remove("strHeadingtext")
            lst.Add(strHeadingText, "strHeadingtext")
        End If
        '
        tbl = MyBase.chpt_Insert_Std(placeBehind, sect, lst, strRptMode, strDivType, 6, False)
        objHFMgr.hf_headers_insert(sect, -1.0)
        '
        If Not showFooter Then
            objHFMgr.hf_footers_delete(sect)
        End If
        '
        'Me.div_edit_Table(tbl, "rptAP")
        'objHFMgr.hf_set_HeaderTagStyle(sect, "tag_appendixPart")
        '                '
        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        objBrndMgr.brnd_recolour_Logo(hf)
        '
        objBrndMgr.brnd_Rebuild_Background(sect, False, False, Me.strTagStyleNameAP)
        Me.chptBase_PageNumbering_Set(sect, False, 1, "div")
        '
        'objHFMgr.hf_set_textColourFooter(sect, RGB(255, 0, 0))
        '
        '
        Return sect
        '
    End Function
    '
    ''' <summary>
    ''' This method will format the Divider page. Appendix and Report Dividers are mostly the same,
    ''' you can specify which you wnat by using strDivTableType (rptAP, or rptBody)
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="strDivTableType"></param>
    Public Sub div_edit_Table(ByRef tbl As Word.Table, Optional strDivTableType As String = "rptBody")
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        Dim strId As String
        Dim fldStyleRef As Word.Field
        '
        If IsNothing(tbl) Then GoTo finis
        '
        Try
            dr = tbl.Rows.Item(2)
            dr.HeightRule = WdRowHeightRule.wdRowHeightAtLeast
            dr.Height = 183.25
            '
            Select Case strDivTableType
                Case "rptAP"
                    'Remove the banner image if it is there
                    drCell = tbl.Range.Cells.Item(1)
                    drCell.Range.ShapeRange.Delete()
                '
                Case "rptBody"
                    'Remove the banner image if it is there
                    drCell = tbl.Range.Cells.Item(1)
                    drCell.Range.ShapeRange.Delete()
                    '
                    dr = tbl.Rows.Item(3)
                    '
                    '*** Get rid of Part xx text in Divider banner
                    GoTo loop0
                    '
                    drCell = dr.Cells.Item(2)
                    dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
                    dr.Height = 11.6
                    dr.Range.Style = Me.objGlobals.glb_get_wrdActiveDoc.Styles("Part xx")
                    '
                    rng = drCell.Range
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    drCell.Range.Text = "Part "
                    '
                    strId = "Part - Number"
                    strId = ControlChars.Quote + strId + ControlChars.Quote
                    '
                    rng = drCell.Range
                    rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                    rng.Move(WdUnits.wdCharacter, -1)

                    fldStyleRef = drCell.Range.Fields.Add(rng, WdFieldType.wdFieldStyleRef, strId, True)
                    'fldPartNum.Update()
                    fldStyleRef.Update()
                    '
                    rng = drCell.Range
                    rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                    rng.Move(WdUnits.wdCharacter, -1)
                    rng.Text = " : "
                    rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                    'rng.Move(WdUnits.wdCharacter, -1)
                    'rng.Text = "Economic Impact"
                    strId = "Part - Heading (Banner)"
                    strId = ControlChars.Quote + strId + ControlChars.Quote
                    fldStyleRef = drCell.Range.Fields.Add(rng, WdFieldType.wdFieldStyleRef, strId, True)
                    fldStyleRef.Update()
loop0:
            End Select

        Catch ex As Exception

        End Try
        '
finis:
        '
    End Sub
    '
    Public Sub div_build_Background(ByRef sect As Word.Section, strTagStyleForHeader As String)
        Dim objBrndMgr As New cBrandMgr()
        '
        Try
            objBrndMgr.brnd_Rebuild_Background(sect, True, False, strTagStyleForHeader)
        Catch ex As Exception
            objBrndMgr.brnd_Rebuild_Background(sect, True, False, "Header")
        End Try
        '
    End Sub
    '

End Class
