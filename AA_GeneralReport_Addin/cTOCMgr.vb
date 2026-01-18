Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports System.IO
Public Class cTOCMgr
    'objGlobals is defined in the inheritance hiearchy cGlossary < cChptBase < cSectionMgr
    '
    Inherits cChptBase

    'Public objGlobals As cGlobals

    Public Sub New()
        MyBase.New()
        'Me.objGlobals = New cGlobals()
    End Sub
    '
    Public Function toc_insert_TOCSection(doTOC As Boolean, placeBehind As Boolean, ByRef sect As Word.Section, strRptMode As String) As Word.Section
        Dim objRptMgr As New cReport()
        Dim objSectMgr As New cSectionMgr()
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objBnrMgr As New cChptBanner()
        Dim indentLeft As Single
        Dim sectNew As Word.Section
        Dim lstOfOldSettings As New Collection()
        Dim myDoc As Word.Document
        '
        indentLeft = 0.0
        sectNew = Nothing
        '
        Try
            'sect = Me.objGlobals.glb_get_wrdSel.Sections.Item(1)
            sectNew = sect
            'rng = Me.objGlobals.glb_get_wrdSelRng()
            myDoc = sect.Range.Document
            '
            'If Not Me.toc_has_TOCSection(myDoc) Then
            '
            Select Case strRptMode
                Case objRptMgr.rpt_isPrt, objRptMgr.modeShort, objRptMgr.rpt_isLnd
                    If strRptMode = objRptMgr.rpt_isPrt Or strRptMode = objRptMgr.modeShort Then
                        sectNew = objSectMgr.sct_insert_Section(placeBehind, sect, 3, "newPage", False, "prt", Me.objGlobals.glb_getDimensions_toc_Prt())
                    End If
                    '
                    If strRptMode = objRptMgr.rpt_isLnd Then
                        sectNew = objSectMgr.sct_insert_Section(placeBehind, sect, 3, "newPage", False, "lnd", Me.objGlobals.glb_getDimensions_toc_Lnd())
                    End If
                    '
                    objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_toc))
                    '
                    'objHFMgr.hf_footers_DeleteContents_All(sectNew)
                    objHFMgr.hf_footers_delete(sectNew)
                    '
                    Me.toc_insert_ContentsHeader(sectNew)
                    '
                    If doTOC Then Me.toc_replace_TOCField(sectNew, "aac_TOC_Levels02")

                Case objRptMgr.rpt_isBrief
                    'Me.toc_replace_TOCField(objGlobals.glb_get_wrdSelRng, "aac_TOC_Levels02")
            End Select
            '
            '
        Catch ex As Exception

        End Try
        '
        Return sectNew
    End Function
    '
    ''' <summary>
    ''' This method will take a basic CHapter section and convert it to a TOC section
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function toc_convert_toTOC(ByRef sect As Word.Section) As Word.Section
        Dim objHFMgr As New cHeaderFooterMgr()
        '
        objHFMgr.hf_footers_delete(sect)
        Me.toc_insert_ContentsHeader(sect)
        '
        Return sect
    End Function

    '
    Public Sub toc_insert_ContentsHeader(ByRef sect As Word.Section)
        Dim hf As Word.HeaderFooter
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim drCell As Word.Cell
        '
        hf = sect.Headers.Item(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary)
        rng = hf.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng.Paragraphs.Add(rng)
        rng.Move(WdUnits.wdParagraph, 1)
        '
        tbl = Me.chptBase_insert_TableAtRange(rng, 1, 1, 56.8, "Header")
        '
        drCell = tbl.Range.Cells.Item(1)
        drCell.Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("TOC Heading")
        drCell.Range.Text = "Contents"
        '
        Me.toc_adjust_ContentsHeader(tbl, 0.0, True)

    End Sub
    '
    '
    ''' <summary>
    ''' This method will return true if the current selection is in the TOC
    ''' Note that there is a duplicate of this function in cFormatMgr.. It's there
    ''' or startup performance reasons only
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function toc_is_TOCSection(ByRef sect As Section) As Boolean
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objBnrMgr As New cChptBanner()
        Dim rslt As Boolean
        Dim objWCAGMr As New cWCAGMgr()
        Dim strTag As String
        '
        rslt = False
        '
        strTag = objHfMgr.hf_tags_getTagStyleName(sect)
        If strTag = objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_toc) Then rslt = True
        '
        'Check for WCAG version
        If objWCAGMr.wcag_docProps_isAccessible() Then

        End If
        '
        '
        Return rslt
    End Function
    '
    ''' <summary>
    ''' THis method will return true if the document myDoc has a TOC Section..If it does return true, then
    ''' sect is set to that section
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function toc_has_TOCSection(ByRef myDoc As Word.Document) As Boolean
        Dim rslt As Boolean
        '
        rslt = False
        For Each sect In myDoc.Sections
            If Me.toc_is_TOCSection(sect) Then
                rslt = True
                Exit For
            End If
        Next

        Return rslt
    End Function
    '
    '
    ''' <summary>
    ''' This method will return nothing if the document does not have a TOC section... If it does it
    ''' will return the section
    ''' </summary>
    ''' <returns></returns>
    Public Function toc_get_TOCSection() As Word.Section
        Dim objBnrMgr As New cChptBanner()
        Dim rslt As Boolean
        Dim objsectMgr As New cSectionMgr()
        Dim strTag As String
        Dim sect As Word.Section
        '
        sect = Nothing
        'strTag = objHfMgr.hf_tags_getTagStyleName(sect)'
        strTag = objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_toc)
        rslt = objsectMgr.sct_has_strTag(objGlobals.glb_get_wrdActiveDoc, strTag, sect)
        '
        Return sect
        '
    End Function


    '
#Region "Styles"
    ''' <summary>
    ''' This is the original form of this method. The other overload allows us
    ''' to force a response with the input of strRprMode as a variable
    ''' </summary>
    Public Sub TOC_Styles_AdjustForReportMode()
        Dim objRptMgr As New cReport()
        Dim strRptMode As String
        '
        strRptMode = objRptMgr.Rpt_Mode_Get()
        Me.TOC_Styles_AdjustForReportMode(strRptMode)
        '
    End Sub


    ''' <summary>
    ''' This method will adjust the right tab stop to suit each of the
    ''' Report Modes
    ''' </summary>
    Public Sub TOC_Styles_AdjustForReportMode(strRptMode As String)
        Dim objRptMgr As New cReport()
        'Dim strRptMode As String
        Dim tabStop, rightIndent As Single
        '
        'strRptMode = objRptMgr.Rpt_Mode_Get()
        '
        Select Case strRptMode
            Case objRptMgr.rpt_isPrt, objRptMgr.modeShort
                tabStop = 382.15
                '
                TOC_Styles_TOCTabStop("TOC 1", tabStop)
                TOC_Styles_TOCTabStop("TOC 2", tabStop)
                TOC_Styles_TOCTabStop("TOC 3", tabStop)
                TOC_Styles_TOCTabStop("TOC 4", tabStop)
                TOC_Styles_TOCTabStop("TOC 5", tabStop)
                TOC_Styles_TOCTabStop("TOC 6", tabStop)
                TOC_Styles_TOCTabStop("TOC 7", tabStop)
                TOC_Styles_TOCTabStop("TOC 8", tabStop)
                TOC_Styles_TOCTabStop("TOC 9", tabStop)
                '
                TOC_Styles_TOCTabStop("Table of Figures", tabStop)
                TOC_Styles_TOCTabStop("Table of Authorities", tabStop)
                '
            Case objRptMgr.rpt_isLnd
                tabStop = 630.0
                rightIndent = 67.4
                '
                TOC_Styles_TOCTabStop("TOC 1", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 2", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 3", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 4", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 5", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 6", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 7", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 8", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 9", tabStop, rightIndent)
                '
                TOC_Styles_TOCTabStop("Table of Figures", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("Table of Authorities", tabStop, rightIndent)

            Case objRptMgr.rpt_isBrief
                'tabStop = 382.15
                tabStop = 432
                rightIndent = 1.3
                '
                TOC_Styles_TOCTabStop("TOC 1", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 2", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 3", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 4", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 5", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 6", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 7", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 8", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 9", tabStop, rightIndent)
                '
                TOC_Styles_TOCTabStop("Table of Figures", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("Table of Authorities", tabStop, rightIndent)
                '

        End Select
    End Sub
    '
    Public Function toc_get_ContentsTable(ByRef sect As Word.Section, autoFitTable As Boolean) As Word.Table
        Dim tbl As Word.Table
        Dim styl, toc2Style As Word.Style
        Dim hf As Word.HeaderFooter
        Dim foundTable As Boolean
        Dim tblWidth As Single
        Dim myDoc As Word.Document
        '
        myDoc = sect.Range.Document
        toc2Style = myDoc.Styles.Item("TOC 2")
        '
        tblWidth = objGlobals.glb_get_widthBetweenMargins(sect) - toc2Style.ParagraphFormat.RightIndent
        '
        tbl = Nothing
        foundTable = False
        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        '
        Try
            If hf.Range.Tables.Count <> 0 Then
                For Each tbl In hf.Range.Tables
                    For Each drCell In tbl.Range.Cells
                        styl = drCell.Range.Style
                        If styl.NameLocal Like "TOC Heading*" Then
                            foundTable = True
                            Exit For
                        End If
                    Next drCell
                    '
                    If foundTable Then
                        If autoFitTable Then
                            'objGlobals.glb_tbls_AutoFitRegularTable(tbl)
                            objGlobals.glb_tbls_AutoFitRegularTableToSize(tbl, tblWidth)
                        End If
                        Exit For
                    End If
                Next tbl
            End If
        Catch ex As Exception
            tbl = Nothing
        End Try

        Return tbl
    End Function
    '
    ''' <summary>
    ''' This method will accept the 'Contents Header' table and adjust its width to either a 
    ''' a value related to the width between margings and the right edge of the TOC 2 style.
    ''' (width between margins - right Indent). Or, if setWidthStyle is false it will be 
    ''' set to the width between the margins.
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="tblWidth"></param>
    ''' <param name="setWidthToStyle"></param>
    Public Sub toc_adjust_ContentsHeader(ByRef tbl As Word.Table, tblWidth As Single, setWidthToStyle As Boolean)
        Dim myDoc As Word.Document
        Dim toc2Style As Word.Style
        Dim sect As Word.Section
        '
        sect = tbl.Range.Sections.Item(1)
        myDoc = sect.Range.Document
        toc2Style = myDoc.Styles.Item("TOC 2")

        tblWidth = objGlobals.glb_get_widthBetweenMargins(sect)
        '
        If setWidthToStyle Then
            tblWidth = objGlobals.glb_get_widthBetweenMargins(sect) - toc2Style.ParagraphFormat.RightIndent
            objGlobals.glb_tbls_AutoFitRegularTableToSize(tbl, tblWidth)
        Else
            objGlobals.glb_tbls_AutoFitRegularTableToSize(tbl, tblWidth)
        End If
        '
        '
    End Sub
    '
    ''' <summary>
    ''' This method will adjust the right tab stop to suit each of the
    ''' Report Modes
    ''' </summary>
    Public Sub toc_Styles_AdjustForOrientation(ByRef sect As Word.Section)
        Dim objRptMgr As New cReport()
        Dim strOrientation As String
        Dim tabStop, rightIndent As Single
        '
        strOrientation = objGlobals.glb_sect_getOrientation(sect)
        '
        Select Case strOrientation
            Case "prt"
                tabStop = 382.15
                '
                TOC_Styles_TOCTabStop("TOC 1", tabStop)
                TOC_Styles_TOCTabStop("TOC 2", tabStop)
                TOC_Styles_TOCTabStop("TOC 3", tabStop)
                TOC_Styles_TOCTabStop("TOC 4", tabStop)
                TOC_Styles_TOCTabStop("TOC 5", tabStop)
                TOC_Styles_TOCTabStop("TOC 6", tabStop)
                TOC_Styles_TOCTabStop("TOC 7", tabStop)
                TOC_Styles_TOCTabStop("TOC 8", tabStop)
                TOC_Styles_TOCTabStop("TOC 9", tabStop)
                '
                TOC_Styles_TOCTabStop("Table of Figures", tabStop)
                TOC_Styles_TOCTabStop("Table of Authorities", tabStop)
                '
            Case "lnd"
                tabStop = 630.0
                'rightIndent = 67.4
                rightIndent = 57.4
                '
                TOC_Styles_TOCTabStop("TOC 1", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 2", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 3", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 4", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 5", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 6", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 7", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 8", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("TOC 9", tabStop, rightIndent)
                '
                TOC_Styles_TOCTabStop("Table of Figures", tabStop, rightIndent)
                TOC_Styles_TOCTabStop("Table of Authorities", tabStop, rightIndent)

        End Select
    End Sub
    '
    '
    ''' <summary>
    ''' This method will set the far right, right aligned tab stop for the specified style
    ''' to a value of tabStop
    ''' </summary>
    ''' <param name="strTOCStyleName"></param>
    ''' <param name="tabStop"></param>
    Private Sub TOC_Styles_TOCTabStop(strTOCStyleName As String, tabStop As Single, Optional rightIndent As Single = 42.55)
        Dim objGlobals As New cGlobals()
        Dim myDoc As Word.Document
        Dim tbStop As Word.TabStop
        Dim tocStyle As Word.Style
        '
        Try
            myDoc = objGlobals.glb_get_wrdActiveDoc()
            tocStyle = myDoc.Styles.Item(strTOCStyleName)
            tocStyle.ParagraphFormat.TabStops.ClearAll()
            tbStop = tocStyle.ParagraphFormat.TabStops.Add(tabStop, WdTabAlignment.wdAlignTabRight)
            '
            tocStyle.ParagraphFormat.RightIndent = rightIndent
        Catch ex As Exception

        End Try
        '
    End Sub
    '
#End Region

#Region "Update"
    '
    Public Sub toc_update_TOCs(ByRef myDoc As Word.Document)
        Dim docTOCs As TablesOfContents
        Dim docTOC As TableOfContents
        Dim i As Integer
        '
        '
        '***AlexR fix 2015.12.17 - for the "sticky" TOC chapter headings issue.
        'this forces a refresh of the "App - Context" and "Chpt - Context" fields before updating the TOC
        'these fields are linked to Heading 1, hidden below the chapter numbers, and picked up by the TOC 4 style (italic purple chapter headings)
        Dim f As Field
        For Each f In myDoc.Fields
            If f.Type = WdFieldType.wdFieldStyleRef Or f.Type = WdFieldType.wdFieldPage Or f.Type = WdFieldType.wdFieldSequence Then
                f.Update() 'Could use the following but not going to: If (f.result.style = "Chpt - Context" Or f.result.style = "App - Context") Then ...
            End If
            '
            'If f.Type = WdFieldType.wdfield
        Next f
        '*** end fix
        '
        '** Why do we have remnants of the Envelope TOC
        docTOCs = myDoc.TablesOfContents
        'docTOC = docTOCs.Item(1)
        'docTOC2 = docTOCs.Item(2)
        '
        'The following is a patch until I can work out what is going on
        '
        For i = 1 To docTOCs.Count
            If i = 1 Then
                'docTOC = docTOCs.Item(i)
                'docTOC.Update()
            End If
        Next

        For Each docTOC In docTOCs
            docTOC.UpdatePageNumbers()
            docTOC.Update()
        Next docTOC
        '

        For Each docTOC In docTOCs
            docTOC.UpdatePageNumbers()
            docTOC.Update()
        Next docTOC
        '
    End Sub
    '
    '
    ''' <summary>
    ''' Update the Table(s) of Figures
    ''' </summary>
    Public Sub toc_upDate_TOFs()
        Dim j As Integer
        Dim myDoc As Word.Document
        '
        myDoc = Me.objGlobals.glb_get_wrdActiveDoc()
        '
        For j = 1 To myDoc.TablesOfFigures.Count
            myDoc.TablesOfFigures(j).Update()
        Next j

    End Sub
    '
#End Region
    '
#Region "TOC Field"
    ''' <summary>
    ''' This method will replace the exsting TOC Field(s) with a 1 level TOC Field, as well
    ''' as Table of Figures/Tables/Boxes. The cursor must be in a TOC Section
    ''' </summary>
    ''' <returns></returns>
    Public Function toc_replace_TOCField_Levels(Optional strNumLevels As String = "aac_TOC_Levels01") As Boolean
        Dim rslt As Boolean
        Dim sect As Word.Section
        Dim objRptMgr As New cReport()
        Dim strRptMode As String
        Dim objGlobals As New cGlobals()
        Dim myDoc As Word.Document
        Dim rng As Word.Range
        Dim fld As Word.Field
        Dim myTOC As Word.TableOfContents


        '
        fld = Nothing
        rslt = False
        sect = Me.objGlobals.glb_get_wrdSelRng.Sections.Item(1)
        myDoc = sect.Range.Document
        strRptMode = objRptMgr.Rpt_Mode_Get()
        '
        Select Case strRptMode
            Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                If Me.toc_is_TOCSection(sect) Then
                    Me.toc_replace_TOCField(sect, strNumLevels)
                    rslt = True
                Else
                    rslt = False
                End If

            Case objRptMgr.rpt_isBrief
                Try

                    myTOC = myDoc.TablesOfContents.Item(1)
                    rng = myTOC.Range
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    'myTOC.
                    'myTOC = myDoc.Fields.Item(WdFieldType.wdFieldTOC)
                    myTOC.Range.Delete()
                    'objGlobals.glb_screen_updateLeaveAsItWas()
                    'myTOC.'
                    'MsgBox("Found TOC Field")
                    'fld.Select()
                    'objGlobals.glb_get_wrdApp.Selection.Delete()
                    rng = Me.toc_replace_TOCField(rng, strNumLevels)
                    rng.Paragraphs.Item(1).Range.Delete()
                    rslt = True

                Catch ex As Exception
                    myTOC = Nothing
                    rslt = False
                End Try
                '
        End Select
        '
        Return rslt
        '
    End Function
    '

    '
    Public Function toc_replace_TOCField(ByRef rng As Word.Range, tocType As String) As Word.Range
        Dim objParas As New cParas()
        Dim objWCAGMgr As New cWCAGMgr()
        Dim objRptMgr As New cReport()
        Dim myDoc As Word.Document
        Dim myTOC As Word.TableOfContents
        Dim fld As Word.Field
        Dim para As Word.Paragraph
        Dim styleHeading_1, styleHeading_2, styleHeading_3 As Word.Style
        Dim styleDividerChpt, styleDividerApp As Word.Style
        Dim styleHeading_1_ES As Word.Style
        Dim styleHeading_1_AP, styleHeading_2_AP, styleHeading_3_AP As Word.Style
        Dim styleHeading_Glossary, styleTOFHeading As Word.Style
        '
        myDoc = rng.Document
        '
        'Me.toc_Styles_AdjustForOrientation(sect)
        '
        'Remove the fields so that there is no clash
        '
        For Each fld In myDoc.Fields
            If fld.Type = WdFieldType.wdFieldTOC Then
                fld.Select()
                rng = objGlobals.glb_get_wrdSelRng()
                fld.Delete()
            End If
        Next
        '

        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Style = myDoc.Styles.Item("Body Text")
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        'Set up the styles we are going to use
        styleHeading_1 = myDoc.Styles.Item("Heading 1")
        styleHeading_2 = myDoc.Styles.Item("Heading 2")
        styleHeading_3 = myDoc.Styles.Item("Heading 3")
        '
        'Part - Heading (Banner)
        'stylePart_xx = myDoc.Styles.Item("Part xx")
        styleDividerChpt = myDoc.Styles.Item("Part - Heading (Banner)")
        '
        styleHeading_1_ES = myDoc.Styles.Item("Heading 1 (ES)")
        styleDividerApp = myDoc.Styles.Item("App - Divider (Heading)")
        '
        If objGlobals._glb_doApp_as_HeadingAP Then
            styleHeading_1_AP = myDoc.Styles.Item("Heading 1 (AP)")
            styleHeading_2_AP = myDoc.Styles.Item("Heading 2 (AP)")
            styleHeading_3_AP = myDoc.Styles.Item("Heading 3 (AP)")
        Else
            styleHeading_1_AP = myDoc.Styles.Item("Heading 6")
            styleHeading_2_AP = myDoc.Styles.Item("Heading 7")
            styleHeading_3_AP = myDoc.Styles.Item("Heading 8")
        End If
        '
        styleHeading_Glossary = myDoc.Styles.Item("Heading (glossary)")
        styleTOFHeading = myDoc.Styles.Item("TOC TOFSubHeading")
        '
        Select Case tocType
            Case "aac_TOC_Levels01"
                '
                myTOC = myDoc.TablesOfContents.Add(rng, False, 1, 1,,,,,, True)
                myTOC.HeadingStyles.Add(styleHeading_1, 2)
                '
                myTOC.HeadingStyles.Add(styleDividerChpt, 1)
                '
                myTOC.HeadingStyles.Add(styleHeading_1_ES, 1)
                myTOC.HeadingStyles.Add(styleDividerApp, 1)
                '
                myTOC.HeadingStyles.Add(styleHeading_1_AP, 2)
                '
                myTOC.HeadingStyles.Add(styleHeading_Glossary, 1)
                '        
                rng = myTOC.Range
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                '

            Case "aac_TOC_Levels02"
                myTOC = myDoc.TablesOfContents.Add(rng, False, 2, 2,,,,,, True)
                myTOC.HeadingStyles.Add(styleHeading_1, 2)
                myTOC.HeadingStyles.Add(styleHeading_2, 3)
                '
                myTOC.HeadingStyles.Add(styleDividerChpt, 1)
                '
                myTOC.HeadingStyles.Add(styleHeading_1_ES, 1)
                myTOC.HeadingStyles.Add(styleDividerApp, 1)
                '
                myTOC.HeadingStyles.Add(styleHeading_1_AP, 2)
                myTOC.HeadingStyles.Add(styleHeading_2_AP, 3)
                '
                myTOC.HeadingStyles.Add(styleHeading_Glossary, 1)
                '        
                rng = myTOC.Range
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                '
            Case "aac_TOC_Levels03"
                myTOC = myDoc.TablesOfContents.Add(rng, False, 2, 3,,,,,, True)
                myTOC.HeadingStyles.Add(styleHeading_1, 2)
                myTOC.HeadingStyles.Add(styleHeading_2, 3)
                myTOC.HeadingStyles.Add(styleHeading_3, 4)
                '
                myTOC.HeadingStyles.Add(styleDividerChpt, 1)
                '
                myTOC.HeadingStyles.Add(styleHeading_1_ES, 1)
                myTOC.HeadingStyles.Add(styleDividerApp, 1)
                '
                myTOC.HeadingStyles.Add(styleHeading_1_AP, 2)
                myTOC.HeadingStyles.Add(styleHeading_2_AP, 3)
                myTOC.HeadingStyles.Add(styleHeading_3_AP, 4)
                '
                myTOC.HeadingStyles.Add(styleHeading_Glossary, 1)
                '
                rng = myTOC.Range
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                '
        End Select
        '
        para = rng.Paragraphs.Add(rng)
        rng = para.Next.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        'rng = Me.insert_TableOfFigures(myDoc, rng)
        'rng = Me.insert_TableOfTables(myDoc, rng)
        'rng = Me.insert_TableOfBoxes(myDoc, rng)
        '
        '
        'Do WCAG 'Contents' Heading
        If objWCAGMgr.wcag_docProps_isAccessible() Then
            'rng = sect.Range
            'rng.Paragraphs.Add(rng)
            'rng.Paragraphs.Add(rng)
            'rng = sect.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'rng.Style = myDoc.Styles.Item("TOC Heading")
            'rng.Text = "Contents"

            '
            'rng = sect.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            'rng.Move(WdUnits.wdCharacter, -1)
        End If
        '
finis:
        '
        Return rng
    End Function

    '
    ''' <summary>
    ''' This method will replace the TOC field with either one or two levels. The tocType is
    ''' "aac_TOC_Levels02" or "aac_TOC_Levels03"... It is like this because it started life
    ''' as something that imported prebuilt TOC's from AutoText.. Due to instabilities in 
    ''' AutoText (20201126) I did this in software
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="tocType"></param>
    ''' <returns></returns>
    Public Function toc_replace_TOCField(ByRef sect As Word.Section, tocType As String) As Word.Range
        Dim objParas As New cParas()
        Dim objWCAGMgr As New cWCAGMgr()
        Dim objRptMgr As New cReport()
        Dim myDoc As Word.Document
        Dim rng As Word.Range
        'Dim myTOC As Word.TableOfContents
        Dim fld As Word.Field
        Dim para As Word.Paragraph
        '
        myDoc = sect.Range.Document
        rng = sect.Range()
        '
        'Me.toc_Styles_AdjustForOrientation(sect)
        '
        'Remove the fields so that there is no clash
        '
        For Each fld In myDoc.Fields
            If fld.Type = WdFieldType.wdFieldTOC Then
                fld.Delete()
            End If
        Next
        '
        objParas.paras_delete_Paragraphs(rng, 3)
        '
        rng = Me.toc_replace_TOCField(rng, tocType)
        '
        'Now add Tables of Figures, Tbales and Boxes
        para = rng.Paragraphs.Add(rng)
        rng = para.Next.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        rng = Me.insert_TableOfFigures(myDoc, rng)
        rng = Me.insert_TableOfTables(myDoc, rng)
        rng = Me.insert_TableOfBoxes(myDoc, rng)
        '
        'Do WCAG 'Contents' Heading
        If objWCAGMgr.wcag_docProps_isAccessible() Then
            rng = sect.Range
            rng.Paragraphs.Add(rng)
            rng.Paragraphs.Add(rng)
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Style = myDoc.Styles.Item("TOC Heading")
            rng.Text = "Contents"

            '
            'rng = sect.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            'rng.Move(WdUnits.wdCharacter, -1)
        End If
        '
        '
        Return rng
    End Function
    '
    Public Function insert_TableOfFigures(myDoc As Word.Document, ByRef rng As Word.Range) As Word.Range
        Dim styleTOFHeading As Word.Style
        Dim para As Word.Paragraph
        Dim fld As Word.Field

        styleTOFHeading = myDoc.Styles.Item("TOC TOFSubHeading")
        rng.Text = "Figures" + vbCrLf
        para = rng.Paragraphs.Item(1)
        para.Range.Style = styleTOFHeading
        '
        '*** Do this if we wnat the Table Of Figures etc to start on a new page
        'para.Format.PageBreakBefore = True
        '***
        '
        para = para.Next
        rng = para.Range
        para.Range.Style = myDoc.Styles.Item("TOC General")
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        fld = myDoc.Fields.Add(rng, WdFieldType.wdFieldTOC, "\h \c " + """Figure ES""")
        fld.Select()
        rng = Me.objGlobals.glb_get_wrdSel.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        fld = myDoc.Fields.Add(rng, WdFieldType.wdFieldTOC, "\h \c " + """Figure""")
        fld.Select()
        rng = Me.objGlobals.glb_get_wrdSel.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        fld = myDoc.Fields.Add(rng, WdFieldType.wdFieldTOC, "\h \c " + """Figure AP""")
        fld.Select()
        rng = Me.objGlobals.glb_get_wrdSel.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
        rng.Text = vbCrLf
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
        Return rng
        '
    End Function
    '
    '
    Public Function insert_TableOfTables(myDoc As Word.Document, ByRef rng As Word.Range) As Word.Range
        Dim styleTOFHeading As Word.Style
        Dim para As Word.Paragraph
        Dim fld As Word.Field

        styleTOFHeading = myDoc.Styles.Item("TOC TOFSubHeading")
        rng.Text = "Tables" + vbCrLf
        para = rng.Paragraphs.Item(1)
        para.Range.Style = styleTOFHeading
        para = para.Next
        rng = para.Range
        para.Range.Style = myDoc.Styles.Item("TOC General")
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        fld = myDoc.Fields.Add(rng, WdFieldType.wdFieldTOC, "\h \c " + """Table ES""")
        fld.Select()
        rng = Me.objGlobals.glb_get_wrdSel.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng.Collapse(WdCollapseDirection.wdCollapseStart)

        fld = myDoc.Fields.Add(rng, WdFieldType.wdFieldTOC, "\h \c " + """Table""")
        fld.Select()
        rng = Me.objGlobals.glb_get_wrdSel.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng.Collapse(WdCollapseDirection.wdCollapseStart)


        fld = myDoc.Fields.Add(rng, WdFieldType.wdFieldTOC, "\h \c " + """Table AP""")
        fld.Select()
        rng = Me.objGlobals.glb_get_wrdSel.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        rng.Text = vbCrLf
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
        Return rng

    End Function
    '
    '
    Public Function insert_TableOfBoxes(myDoc As Word.Document, ByRef rng As Word.Range) As Word.Range
        Dim styleTOFHeading As Word.Style
        Dim para As Word.Paragraph
        Dim fld As Word.Field

        styleTOFHeading = myDoc.Styles.Item("TOC TOFSubHeading")
        rng.Text = "Boxes" + vbCrLf
        para = rng.Paragraphs.Item(1)
        para.Range.Style = styleTOFHeading
        para = para.Next
        rng = para.Range
        para.Range.Style = myDoc.Styles.Item("TOC General")
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        fld = myDoc.Fields.Add(rng, WdFieldType.wdFieldTOC, "\h \c " + """Box ES""")
        fld.Select()
        rng = Globals.ThisAddin.Application.Selection.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng.Collapse(WdCollapseDirection.wdCollapseStart)

        fld = myDoc.Fields.Add(rng, WdFieldType.wdFieldTOC, "\h \c " + """Box""")
        fld.Select()
        rng = Globals.ThisAddin.Application.Selection.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng.Collapse(WdCollapseDirection.wdCollapseStart)

        fld = myDoc.Fields.Add(rng, WdFieldType.wdFieldTOC, "\h \c " + """Box AP""")
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        fld.Select()
        rng = Globals.ThisAddin.Application.Selection.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        'rng.Text = vbCrLf
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
        Return rng
        '
    End Function
    '


#End Region

    ''' <summary>
    ''' This method will move the selection to the bottom of the TOC
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub toc_Selection_MoveToTOC(ByRef myDoc As Word.Document)
        Dim myTOC As Word.TableOfContents
        Dim rng As Word.Range

        Try
            'Because we clicked on this let's move the selection
            myTOC = myDoc.TablesOfContents.Item(1)
            rng = myTOC.Range
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            rng.Select()
            '
        Catch ex As Exception

        End Try
    End Sub
End Class
