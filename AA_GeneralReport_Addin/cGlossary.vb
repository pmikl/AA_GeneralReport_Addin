Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cGlossary
    'objGlobals is defined in the inheritance hiearchy cGlossary < cChptBase < cSectionMgr
    '
    Inherits cChptBase
    '
    'Public objGlobals As cGlobals

    Public Sub New()
        MyBase.New()
        '
        'Me.objGlobals = New cGlobals()

    End Sub
    '
    Public Function glos_is_Glossary(ByRef sect As Word.Section) As Boolean
        Dim rslt As Boolean
        Dim objTagsMgr As New cTagsMgr()
        Dim strTagStyleName As String
        Dim objBnrMgr As New cChptBanner()
        '
        rslt = False
        '
        strTagStyleName = objTagsMgr.tags_get_tagStyleName(sect)
        '
        If strTagStyleName = objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_glos) Then rslt = True

        Return rslt
        '
    End Function
    '
    Public Function glos_insert_Glossary(placeBehind As Boolean, ByRef sect As Word.Section) As Word.Range
        Dim strSectType As String
        Dim objBnrMgr As New cChptBanner()
        Dim objRptMgr As New cReport()
        Dim objPara As New cParas()
        Dim strRptMode As String
        Dim tbl As Word.Table
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        '
        strRptMode = objRptMgr.Rpt_Mode_Get()
        '
        rng = Nothing
        tbl = Nothing
        strSectType = objBnrMgr.tag_glos
        '
        Select Case strRptMode
            Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                rng = Me.glos_insert_General(placeBehind, sect, strSectType)
            Case objRptMgr.rpt_isBrief
                '
                'MsgBox("Glossary")
                rng = objGlobals.glb_get_wrdSelRng
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                '
                rng = objPara.paras_insert_numParas(rng, 1)
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                objPara.paras_add_textAndStyle(rng, "Glossary", "Heading (glossary)")
                rng.Move(WdUnits.wdParagraph, 1)
                '
                tbl = Me.glos_Insert_TableForGlossary(rng)
                drCell = tbl.Range.Cells.Item(3)
                rng = drCell.Range
                rng.MoveEnd(WdUnits.wdCharacter, -1)
                '
                rng.Select()

        End Select
        '
        Return rng
    End Function
    '
    ''' <summary>
    ''' This function will select the contents of the first cell of a Glossary Table... It will also return the
    ''' range of that selection (independently). If there was no Table then the range is returned as Nothing
    ''' the range of that 
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function glos_select_GlossaryFirstEntry(ByRef sect As Word.Section) As Word.Range
        Dim tbl As Word.Table
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        '
        rng = Nothing
        If Not (sect.Range.Tables.Count = 0) Then
            tbl = sect.Range.Tables.Item(1)
            drCell = tbl.Range.Cells.Item(3)
            rng = drCell.Range
            rng.MoveEnd(WdUnits.wdCharacter, -1)
            rng.Select()
            '
        End If
        '
        Return rng
    End Function

    '
    Public Function glos_insert_Biblio(placeBehind As Boolean, ByRef sect As Word.Section) As Word.Range
        Dim strSectType As String
        Dim objBnrMgr As New cChptBanner()
        Dim objRptMgr As New cReport()
        Dim objPara As New cParas()
        Dim strRptMode As String
        Dim rng As Word.Range
        Dim fld As Word.Field

        strRptMode = objRptMgr.Rpt_Mode_Get()
        rng = Nothing
        '
        strSectType = objBnrMgr.tag_glos_bib
        '
        Select Case strRptMode
            Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                rng = Me.glos_insert_General(placeBehind, sect, strSectType)
                chptBase_PageNumbering_Set(sect, False, 1, "std")
                '
            Case objRptMgr.rpt_isBrief
                rng = objGlobals.glb_get_wrdSelRng
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                '
                rng = objPara.paras_insert_numParas(rng, 1)
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                objPara.paras_add_textAndStyle(rng, "Bibliography", "Heading (glossary)")
                rng.Move(WdUnits.wdParagraph, 1)
                '
                fld = sect.Range.Document.Fields.Add(rng, WdFieldType.wdFieldBibliography,, True)
                fld.Select()
                '
                rng = objGlobals.glb_get_wrdSelRngAll()

                '
                'tbl = Me.glos_Insert_TableForGlossary(rng)
                'drCell = tbl.Range.Cells.Item(3)
                'rng = drCell.Range
                'rng.MoveEnd(WdUnits.wdCharacter, -1)
                '
                rng.Select()

        End Select

        'chptBase_PageNumbering_Set(sect, False, 1, "div")
        '
        Return rng
    End Function

    '
    Public Function glos_insert_Refs(placeBehind As Boolean, ByRef sect As Word.Section) As Word.Range
        Dim strSectType As String
        Dim objBnrMgr As New cChptBanner()
        Dim objRptMgr As New cReport()
        Dim objPara As New cParas()
        Dim strRptMode As String
        Dim rng As Word.Range
        Dim fld As Word.Field

        strRptMode = objRptMgr.Rpt_Mode_Get()
        rng = Nothing
        '
        strSectType = objBnrMgr.tag_glos_refsCited
        '
        Select Case strRptMode
            Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                rng = Me.glos_insert_General(placeBehind, sect, strSectType)
                chptBase_PageNumbering_Set(sect, False, 1, "std")
                '
            Case objRptMgr.rpt_isBrief
                rng = objGlobals.glb_get_wrdSelRng
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                '
                rng = objPara.paras_insert_numParas(rng, 1)
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                objPara.paras_add_textAndStyle(rng, "References", "Heading (glossary)")
                rng.Move(WdUnits.wdParagraph, 1)
                '
                fld = sect.Range.Document.Fields.Add(rng, WdFieldType.wdFieldBibliography,, True)
                fld.Select()
                '
                rng = objGlobals.glb_get_wrdSelRngAll()

                '
                'tbl = Me.glos_Insert_TableForGlossary(rng)
                'drCell = tbl.Range.Cells.Item(3)
                'rng = drCell.Range
                'rng.MoveEnd(WdUnits.wdCharacter, -1)
                '
                rng.Select()

        End Select
        '
        'chptBase_PageNumbering_Set(sect, False, 1, "div")
        '
        Return rng
    End Function
    '
    Public Function glos_insert_WorksCited(placeBehind As Boolean, ByRef sect As Word.Section) As Word.Range
        Dim strSectType As String
        Dim objBnrMgr As New cChptBanner()
        Dim objRptMgr As New cReport()
        Dim objPara As New cParas()
        Dim strRptMode As String
        Dim rng As Word.Range
        Dim fld As Word.Field

        strRptMode = objRptMgr.Rpt_Mode_Get()
        rng = Nothing
        '
        strSectType = objBnrMgr.tag_glos_wrks
        '
        Select Case strRptMode
            Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                rng = Me.glos_insert_General(placeBehind, sect, strSectType)
                chptBase_PageNumbering_Set(sect, False, 1, "std")

            Case objRptMgr.rpt_isBrief
                rng = objGlobals.glb_get_wrdSelRng
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                '
                rng = objPara.paras_insert_numParas(rng, 1)
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                objPara.paras_add_textAndStyle(rng, "Works Cited", "Heading (glossary)")
                rng.Move(WdUnits.wdParagraph, 1)
                '
                fld = sect.Range.Document.Fields.Add(rng, WdFieldType.wdFieldBibliography,, True)
                fld.Select()
                '
                rng = objGlobals.glb_get_wrdSelRngAll()
                '
                rng.Select()

        End Select
        'chptBase_PageNumbering_Set(sect, False, 1, "div")
        '
        Return rng
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will insert a 'Glossary', or  ('Bibliography', 'References' or 'Works Cited'). In either Landscape ('lnd')
    ''' or Portrait ('prt') depending on the Report mode. The Header Table is tagged with the appropriate cChptBanner.bnr_get_tagStyles
    ''' placeBehind =   'not used'
    ''' sect        =   A byref return variable identifying the section created
    ''' strSectType =   'objBnrMgr.sectType_glos' or 'objBnrMgr.sectType_bib, objBnrMgr.sectType_refs, objBnrMgr.sectType_wrks'
    ''' </summary>
    ''' <param name="placeBehind"></param>
    ''' <param name="sect"></param>
    ''' <param name="strSectType"></param>
    ''' <returns></returns>
    Public Function glos_insert_General(placeBehind As Boolean, ByRef sect As Word.Section, strSectType As String) As Word.Range
        Dim objRptMgr As New cReport()
        Dim objBnrMgr As New cChptBanner()
        Dim objFldsMgr As New cFieldsMgr()
        Dim objColsMgr As New cColsHandler()
        Dim objTOCMgr As New cTOCMgr()
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim strTagStyle, strOrientation As String
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim drCell As Word.Cell
        Dim lst As Collection
        Dim strRptMode As String
        Dim fld As Word.Field
        Dim myDoc As Word.Document
        '
        strRptMode = objRptMgr.Rpt_Mode_Get()
        '
        myDoc = sect.Range.Document
        strOrientation = "prt"
        tbl = Nothing
        rng = Nothing
        '
        If strRptMode = objRptMgr.rpt_isLnd Then strOrientation = "lnd"
        '
        'lst = objBnrMgr.bnr_get_BannerSettings(objBnrMgr.sectType_glos, True)
        lst = objBnrMgr.bnr_get_BannerSettings(strSectType, True)
        strTagStyle = CStr(lst.Item("strTagStyle"))
        tbl = MyBase.chpt_Insert_Std(placeBehind, sect, lst, strRptMode, strOrientation)
        '
        'If strOrientation = "lnd" Then objColsMgr.cols_setup_columnStructure(sect, "2_columns")
        'sect.PageSetup.TextColumns.SetCount(2
        '
        objHfMgr.hf_tags_setTagStyle(sect, strTagStyle)
        'chptBase_set_tagStyleInHeaderTable(sect, strTagStyle)
        '
        Select Case strSectType
            Case objBnrMgr.tag_glos
                '
                rng = sect.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                '
                'rng.Move(WdUnits.wdParagraph, 2)
                rng.Move(WdUnits.wdParagraph, 1)
                '
                tbl = Me.glos_Insert_TableForGlossary(rng)
                drCell = tbl.Range.Cells.Item(3)
                rng = drCell.Range
                rng.MoveEnd(WdUnits.wdCharacter, -1)
                '
                rng.Select()
                Me.chptBase_PageNumbering_Set(tbl.Range.Sections.Item(1), True, 1, "es")
                        '
            Case objBnrMgr.tag_glos_bib, objBnrMgr.tag_glos_refsCited, objBnrMgr.tag_glos_wrks
                '
                rng = sect.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Move(WdUnits.wdParagraph, 2)
                fld = sect.Range.Document.Fields.Add(rng, WdFieldType.wdFieldBibliography,, True)
                fld.Select()
                '
                rng = objGlobals.glb_get_wrdSelRngAll()
                '
                'Me.chptBase_PageNumbering_Set(tbl.Range.Sections.Item(1), False, 1, "div")
                Me.chptBase_PageNumbering_Set(sect, False, 1, "std")

                'Me.chptBase_PageNumbering_Set(sect, False, 1, "div")

        End Select
        '
        objFldsMgr.updateSequenceNumbers_Chapters()
        objTOCMgr.toc_update_TOCs(myDoc)
        '
        Return rng
        '
    End Function

    Public Function xglos_insert_Biblio(placeBehind As Boolean, ByRef sect As Word.Section) As Word.Table
        Dim objRptMgr As New cReport()
        Dim objBnrMgr As New cChptBanner()
        Dim objFldsMgr As New cFieldsMgr()
        Dim objColsMgr As cColsHandler
        Dim fld As Word.Field
        Dim rng As Word.Range
        Dim tbl, tbl2 As Word.Table
        Dim lst As New Collection()
        Dim strRptMode As String
        Dim sect2 As Word.Section
        '
        strRptMode = objRptMgr.Rpt_Mode_Get()
        tbl = Nothing
        '
        Select Case strRptMode
            Case objRptMgr.rpt_isPrt, objRptMgr.modeShort
                lst = objBnrMgr.bnr_get_BannerSettings(objBnrMgr.tag_glos_bib, True)
                tbl = MyBase.chpt_Insert_Std(placeBehind, sect, lst, strRptMode)
                '
                rng = objGlobals.glb_set_wrdSel(tbl)
                fld = sect.Range.Document.Fields.Add(rng, WdFieldType.wdFieldBibliography,, True)
                fld.Select()
                '
                Me.chptBase_PageNumbering_Set(tbl.Range.Sections.Item(1), False, 1, "std")
                '
            Case objRptMgr.rpt_isLnd
                objColsMgr = New cColsHandler()
                lst = objBnrMgr.bnr_get_BannerSettings(objBnrMgr.tag_glos_bib, False)
                tbl = MyBase.chpt_Insert_LandscapeReport(placeBehind, sect, lst)
                '
                sect2 = tbl.Range.Sections.Item(1)
                sect2 = tbl.Range.Document.Sections.Item(sect2.Index + 1)
                '
                objColsMgr.cols_setup_columnStructure(sect2, "2_columns")
                rng = sect2.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                tbl2 = Me.glos_Insert_TableForGlossary(rng)

                Me.chptBase_PageNumbering_Set(tbl.Range.Sections.Item(1), False, 1, "std")
                Me.chptBase_PageNumbering_Set(sect, False, 1, "std")
                '
        End Select
        '
        'objFldsMgr.updateSequenceNumbers_Chapters()


        Return tbl
        '
        '
    End Function
    '
    '
#Region "Tables"
    '
    ''' <summary>
    ''' This method will insert the Glossary Table at the Range rng
    ''' </summary>
    ''' <param name="rng"></param>
    Public Function glos_Insert_TableForGlossary(ByRef rng As Word.Range) As Word.Table
        Dim tbl As Word.Table
        Dim tblWidth As Single
        Dim dr As Word.Row
        Dim drCol As Word.Column
        Dim drCell As Word.Cell
        Dim objTools As New cTools()
        Dim objTblMgr As New cTablesMgr()
        Dim sect As Word.Section
        '
        sect = rng.Sections.Item(1)
        If sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait Then
            tbl = rng.Tables.Add(rng, 30, 2)
        Else
            tbl = rng.Tables.Add(rng, 20, 2)
        End If
        '
        tbl.Style = objTblMgr.var_tbl_TableStyleDefault
        '
        tbl.Range.Style = objGlobals.glb_get_wrdActiveDoc().Styles("Glossary")
        tbl.Borders.Enable = False
        tbl.LeftPadding = 0.0
        tbl.TopPadding = 0.0
        tbl.RightPadding = 0.0
        tbl.BottomPadding = 0.0
        '
        dr = tbl.Rows.Item(1)
        'To repeat heading row
        dr.HeadingFormat = True
        '
        tblWidth = tbl.Range.Columns.Item(1).Width + tbl.Range.Columns.Item(2).Width
        drCol = tbl.Range.Columns.Item(1)
        drCol.Width = 90.0
        drCol = tbl.Range.Columns.Item(2)
        drCol.Width = tblWidth - 90.0
        '
        '
        '
        drCell = tbl.Range.Cells.Item(1)
        drCell.Range.Text = "Abbreviations"
        drCell.Range.Font.Color = WdColor.wdColorWhite
        drCell.Range.Font.Bold = True
        '
        drCell = tbl.Range.Cells.Item(2)
        drCell.Range.Text = "Definitions"
        drCell.Range.Font.Color = WdColor.wdColorWhite
        drCell.Range.Font.Bold = True
        '
        For i = 3 To 4
            drCell = tbl.Range.Cells.Item(i)
            '
            If objTools.tools_math_isOdd(i) Then
                drCell.Range.Text = "Overtype here"
            Else
                drCell.Range.Text = "Type here using the 'Glossary' style from the 'Styles' tab"
            End If

        Next
        '
        Return tbl

    End Function

    '
    ''' <summary>
    ''' This method expects as input a Banner Table.. it will insert a "Glossary" table
    ''' directly below that table
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub glos_Insert_TableForGlossary(ByRef sect As Word.Section)
        Dim rng As Word.Range
        Dim tbl As Word.Table
        '
        rng = sect.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        'If we have a banner Table insert just below this, otherwise do it at the top of the page
        If rng.Tables.Count <> 0 Then
            tbl = rng.Tables.Item(1)
            rng = tbl.Range
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            rng.Move(WdUnits.wdParagraph, 1)
            Me.glos_Insert_TableForGlossary(rng)
        Else
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            Me.glos_Insert_TableForGlossary(rng)
            '
        End If
        '
    End Sub
#End Region




End Class
