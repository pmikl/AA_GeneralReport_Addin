Imports System.IO
Imports System.Collections
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core

Public Class cWCAGMgr
    Public _wcag_color_backcolour_old As Long       'This is the original backcolor
    Public _wcag_backcolour As Long                 'This is the alternate fill colour for wcag documents
    Public _wcag_color_heading_purple As Long       'Light purple for sub headings
    Public _wcag_color_backcolour_tblHeader As Long 'This is the header row back colour
    '
    Public _wcag_tables_leftIndent_Wide             'Left indent for wide tables
    '
    Public flg_killConversion As Int32              'If set to 1, then we'll exit the loop in convertDocs
    Public currentDoc As Word.Document
    Public _saveDirectoryFullName As String         'The directory into which the converted file is saved

    Public _strFileToConvert As String               'Full name of the file to convert
    Public _strFileTag As String
    Public _convertError As Boolean
    '
    Public Sub New()
        Me.flg_killConversion = 0
        Me.currentDoc = Nothing
        Me._strFileToConvert = ""
        Me._saveDirectoryFullName = ""
        '
        Me._wcag_color_backcolour_old = RGB(20, 0, 52)
        'Me._wcag_backcolour = RGB(196, 172, 221)
        Me._wcag_backcolour = RGB(255, 255, 255)
        Me._wcag_color_backcolour_tblHeader = RGB(20, 0, 52)
        Me._wcag_color_heading_purple = RGB(108, 62, 153)
        '
        Me._wcag_tables_leftIndent_Wide = -98.15
        '
    End Sub
    '
    Public Sub New(strFileToConvertFullName As String, strSaveDirectory As String, fileTag As String)
        Me.flg_killConversion = 0
        Me.currentDoc = Nothing
        Me._strFileToConvert = strFileToConvertFullName
        Me._saveDirectoryFullName = strSaveDirectory
        Me._strFileTag = fileTag
        '
        Me._convertError = False
        '
        Me._wcag_color_backcolour_old = RGB(20, 0, 52)
        Me._wcag_backcolour = RGB(196, 172, 221)
        Me._wcag_color_backcolour_tblHeader = RGB(20, 0, 52)
        Me._wcag_color_heading_purple = RGB(108, 62, 153)
        '
        Me._wcag_tables_leftIndent_Wide = -98.15
    End Sub
    '
    '
    ''' <summary>
    ''' This method will set the Shape (shp) Decorative property to true or False
    ''' </summary>
    ''' <param name="shp"></param>
    ''' <param name="setToTrue"></param>
    Public Sub wcag_set_decorative(ByRef shp As Word.Shape, setToTrue As Boolean)
        '
        Dim objShp As Object
        objShp = shp
        '
        objShp.Decorative = 0
        '
        If setToTrue Then
            objShp.Decorative = 1
        Else
            objShp.Decorative = 0
        End If
        '
    End Sub
    '
    ''' <summary>
    ''' This method will set the InLine Shape (ilshp) Decorative property to true or False
    ''' </summary>
    ''' <param name="ilshp"></param>
    ''' <param name="setToTrue"></param>
    Public Sub wcag_set_decorative(ByRef ilshp As Word.InlineShape, setToTrue As Boolean)
        '
        Dim objShp As Object
        objShp = ilshp
        '
        If setToTrue Then
            objShp.Decorative = 1
        Else
            objShp.Decorative = 0
        End If
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will set the Decorative Property of the first shape in the
    ''' ShapeRange of rng
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="setToTrue"></param>
    Public Sub wcag_set_decorative(ByRef rng As Word.Range, setToTrue As Boolean)
        Dim shp As Word.Shape
        '
        If rng.ShapeRange.Count <> 0 Then
            shp = rng.ShapeRange.Item(1)
            Me.wcag_set_decorative(shp, setToTrue)
        End If
        '
    End Sub

    '
    '
    ''' <summary>
    ''' This method will open, clone and modify to WCAG compliance the document identified by the
    ''' class property 'strFileToConvert' (file fullname). If there was an error on conversion
    ''' </summary>
    Public Sub convertSelectedDoc()
        Dim newDoc As Word.Document
        Dim objFileMgr As New cFileHandler()
        Dim strNewFileName As String
        Dim info As FileInfo
        'Dim evt As eve
        '
        newDoc = Nothing

        info = My.Computer.FileSystem.GetFileInfo(Me._strFileToConvert)
        '
        Try
            If objFileMgr.isFileOpen(info) Then
                Me.currentDoc = Globals.ThisAddin.Application.Documents.Item(Me._strFileToConvert)
                'strNewFileName = Me.getNewFileName(Me.currentDoc, Me._saveDirectoryFullName, Me._strFileTag)
                strNewFileName = objFileMgr.getNewFileName(Me.currentDoc, Me._saveDirectoryFullName, Me._strFileTag)
                newDoc = objFileMgr.cloneToSpecificPath(Me.currentDoc, strNewFileName)
                Me.wcag_doc_ToWCAG(newDoc)
                '
                newDoc.Save()
                newDoc.Close()
                Me._convertError = False
            Else
                'File is not open
                Me.currentDoc = Globals.ThisAddin.Application.Documents.Open(FileName:=Me._strFileToConvert, AddToRecentFiles:=False, Visible:=False)
                'strNewFileName = Me.getNewFileName(Me.currentDoc, Me._saveDirectoryFullName, Me._strFileTag)
                strNewFileName = objFileMgr.getNewFileName(Me.currentDoc, Me._saveDirectoryFullName, Me._strFileTag)

                newDoc = objFileMgr.cloneToSpecificPath(Me.currentDoc, strNewFileName)
                Me.currentDoc.Saved = True
                Me.currentDoc.Close()
                Me.wcag_doc_ToWCAG(newDoc)
                '
                newDoc.Save()
                newDoc.Close()
                '
                Me._convertError = False
            End If
            '
        Catch ex As Exception
            Me._convertError = True
        End Try
        '
    End Sub

    '
    Public Function x_getNewFileName(ByRef sourceDoc As Word.Document, destinationFolderFullName As String, strFileId As String) As String
        Dim objFileMgr As cFileHandler
        Dim strExtension, strNewFileName, strTimeSTamp As String
        Dim tokens() As String
        'Dim info As FileInfo
        '
        'info = My.Computer.FileSystem.GetFileInfo(strFullName)
        objFileMgr = New cFileHandler()
        strTimeSTamp = objFileMgr.file_get_TimeStamp()
        '
        strExtension = Path.GetExtension(sourceDoc.FullName)
        tokens = sourceDoc.Name.Split(".")
        'Now add an id to the file name (eg.g. wcag, dotNet etc)
        If strFileId = "" Then
            strNewFileName = tokens(0) + "-" + strTimeSTamp + strExtension
        Else
            strNewFileName = tokens(0) + "-" + strFileId + "-" + strTimeSTamp + strExtension
        End If
        '
        strNewFileName = destinationFolderFullName + "\" + strNewFileName
        '
        Return strNewFileName
    End Function

    '
    Public Sub xwcag_doc_ToWCAG(ByRef myDoc As Word.Document)
        Call Me.wcag_styles_setForWCAG(myDoc)
        Call Me.doTables(myDoc)
        Call Me.doTextBoxes(myDoc)
    End Sub
    '
    Public Sub wcag_doc_ToWCAG(ByRef myDoc As Word.Document)
        Me.wcag_doc_ToWCAG_entry(False, myDoc)
        Dim objHFMgr As New cHeaderFooterMgr()
        'D'im objFldsMgr As New cFieldsMgr()
        'Dim objGlobals As New cGlobals()
        '
        '
        'objGlobals.glb_screen_update()
        'myDoc.Fields.Update()
        '
        'Me.wcag_convert_backColour(myDoc)
        'objGlobals.glb_screen_update()
        '
        'objFldsMgr.flds_tocs_unlink(myDoc, False)
        '
        'Me.wcag_convert_headersToText(myDoc)
        'Me.wcag_convert_footersToText(myDoc)
        'objFldsMgr.flds_footer_unlink(myDoc)
        '
        '
        'Call Me.wcag_styles_setForWCAG(myDoc)
        '
        'The order matters, otherwise fields may lose their 'result' because sections of document
        'that they rely on have been changed
        '
        '
        '
        'objFldsMgr.flds_footer_unlink(myDoc)
        '
        'Me.wcag_convert_bannersToWCAG(myDoc)
        'myDoc.Fields.Unlink()
        'objFldsMgr.flds_body_unLink(myDoc)
        '
        'objGlobals.glb_screen_update()
        'Tested OK 20220404
        'Me.wcag_docProps_setAccessibity(True, myDoc)
        '
        'Tested OK 20220404
        'Me.wcag_rbn_del(myDoc)

        'objGlobals.glb_screen_update(True)

        'Call Me.doTables(myDoc)
        'Call Me.doTextBoxes(myDoc)
    End Sub

    '
    Private Sub doTextBoxes(ByRef myDoc As Word.Document)
        Dim shp As Word.Shape
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim drCell As Word.Cell
        Dim strText As String
        '
        Try
            For Each shp In myDoc.Shapes
                'obj = TryCast(shp, Microsoft.Office.Tools.Word.Controls.TextBox)
                If shp.Type = MsoShapeType.msoTextBox Then
                    'Found a TextBox (not inline)
                    rng = shp.TextFrame.TextRange
                    If rng.Tables.Count <> 0 Then
                        tbl = rng.Tables.Item(1)
                        drCell = tbl.Range.Cells.Item(1)
                        strText = drCell.Range.Text.Trim()
                        tbl.Title = strText
                        tbl.Descr = "Picture Pullout"
                    End If
                End If
            Next
        Catch ex As Exception
            'MessageBox.Show("fault")
        End Try

    End Sub
    '
    Private Sub doTables(ByRef myDoc As Word.Document)
        Dim objSectMgr As New cSectionMgr()
        Dim objTags As New cTagsMgr()
        Dim strTag, strTagType As String
        Dim tbl As Word.Table
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        Dim strText, strFieldText As String
        Dim styl As Word.Style
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        Dim objCpMgr As New cCoverPageMgr()
        Dim rowNum As Integer
        '
        For Each tbl In myDoc.Tables
            strTag = objTags.tags_get_tagStyleName(tbl)
            strTagType = objTags.getTagType(strTag)
            Select Case strTagType
                Case "coverPage", "contactsPage-Front", "contactsPage-Back", "toc"
                    If strTagType = "coverPage" Then
                        strText = objCpMgr.getCoverPageTitleText(tbl)
                        strText = strText.Replace(vbCr, " ").Replace(vbLf, "")
                        tbl.Title = strText
                        tbl.Descr = "Copyright Acil Allen"
                    End If
                    If strTagType = "contactsPage-Front" Then
                        strText = "Contacts Page - Front"
                        tbl.Title = strText
                        tbl.Descr = "Acil Allen"
                    End If
                    If strTagType = "contactsPage-Back" Then
                        strText = "Contacts Page - Back"
                        tbl.Title = strText
                        tbl.Descr = "Acil Allen"
                    End If
                Case "partBanner", "appendixPart", "execBanner", "chapterBanner", "appendixChapter"
                    '
                    strText = Me.getBannerText(tbl)
                    strFieldText = Me.getSequenceFieldText(tbl)
                    tbl.Title = strText
                    '
                    If strTagType = "partBanner" Or strTagType = "chapterBanner" Or strTagType = "appendixChapter" Then
                        If strFieldText <> "" Then
                            Select Case strTagType
                                Case "partBanner"
                                    tbl.Descr = "Part " + strFieldText
                                    tbl.Title = "Part " + strFieldText + " - " + strText
                                Case "chapterBanner"
                                    tbl.Descr = "Chapter " + strFieldText
                                    tbl.Title = "Chapter " + strFieldText + " - " + strText
                                Case "appendixChapter"
                                    tbl.Descr = "Appendix " + strFieldText
                                    tbl.Title = "Appendix " + strFieldText + " - " + strText
                            End Select
                        End If
                    End If
                Case Else
                    'For all Boxes, Tables, Figures etc
                    drCell = tbl.Range.Cells.Item(1)
                    styl = drCell.Range.Style
                    '
                    'Embedded captions may be in row 1 or 2, so we need to look
                    'in both
                    rowNum = Me.testForEmbeddedCaption(tbl)
                    '
                    If rowNum >= 1 Then
                        dr = tbl.Rows.Item(rowNum)
                        drCell = dr.Cells.Item(1)
                        strText = drCell.Range.Text.Trim()
                        strText = strText.Replace(vbTab, " - ")
                        tbl.Title = strText
                    Else
                        'For table structures with Captions directly above the structure,
                        'but NOT in the structure
                        drCell = tbl.Range.Cells.Item(1)
                        rng = drCell.Range
                        rng.Collapse(WdCollapseDirection.wdCollapseStart)
                        rng.Move(WdUnits.wdParagraph, -1)
                        para = rng.Paragraphs.Item(1)
                        strText = para.Range.Text
                        strText = strText.Replace(vbTab, " - ")
                        tbl.Title = strText
                    End If

            End Select
        Next
        '
    End Sub
    '
    ''' <summary>
    ''' This method will test the first two rows of a table structure to determine if
    ''' either row contains a caption.. It will return the row number (1,2) if a caption
    ''' is found and -1 if one is not found
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function testForEmbeddedCaption(ByRef tbl As Word.Table) As Integer
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        Dim rslt As Integer
        Dim styl As Word.Style
        Dim i As Integer
        '
        rslt = -1
        Try
            For i = 1 To 2 Step 1
                dr = tbl.Rows.Item(i)
                drCell = dr.Cells.Item(1)
                styl = drCell.Range.Style
                If styl.NameLocal Like "Caption*" Then
                    rslt = i
                    Exit For
                End If
            Next i
        Catch ex As Exception
            rslt = -1
        End Try
        '
        Return rslt
    End Function
    '
    Public Function getBannerText(ByRef tbl As Word.Table) As String
        Dim objSectMgr As New cSectionMgr()
        Dim objTags As New cTagsMgr()
        Dim drCell As Word.Cell
        Dim strText, strTag, strTagType As String

        strTag = objTags.tags_get_tagStyleName(tbl)
        strTagType = objTags.getTagType(strTag)
        '
        drCell = tbl.Range.Cells.Item(3)
        strText = drCell.Range.Text.Trim()
        strText = strText.Replace(vbCr, " ").Replace(vbLf, "")
        '
        getBannerText = strText

    End Function
    '
    Public Function getSequenceFieldText(ByRef tbl As Word.Table) As String
        Dim drCell As Word.Cell
        'Dim fld As Word.Field
        Dim strText As String
        '
        strText = ""
        Try
            drCell = tbl.Range.Cells.Item(4)
            If drCell.Range.Fields.Count > 0 Then
                strText = drCell.Range.Fields.Item(1).Result.Text
                strText = strText.Trim()
            End If
        Catch ex As Exception
            strText = ""
        End Try
        '
        Return strText
    End Function
    '
    '
    ''' <summary>
    ''' This method will convert the selected document to comply (partially) with
    ''' WCAG requirements
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub wcag_styles_setForWCAG(ByRef myDoc As Word.Document)
        Dim styl As Word.Style
        Dim objGlobals As New cGlobals()
        '

        'styl = myDoc.Styles.Item("Normal")
        'styl.Font.Size = 9.5
        'styl.Font.Name = "Arial"
        'styl.Font.Color = RGB(0, 0, 0)
        '
        'ByPass the styles conversion if this document has already been through the
        'WCAG conversion
        'If Me.wcag_docProps_isAccessible() Then Exit Sub
        '
        For Each styl In myDoc.Styles
            styl.Font.Color = RGB(0, 0, 0)
            'If styl.Font.Name <> "Arial" Then styl.Font.Name = "Arial"
        Next
        '
        '
        GoTo finis

        For Each styl In myDoc.Styles
            'styl.Font.Color = RGB(0, 0, 0)
            'If styl.Font.Name <> "Arial" Then styl.Font.Name = "Arial"
            '
            GoTo finis
            '
            Select Case styl.NameLocal
                Case "Normal"
                    'styl.Font.Size = 11
                    'styl.Font.Color = RGB(0, 0, 0)
                    'styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast
                    'styl.ParagraphFormat.LineSpacing = 16
                Case "LetterSubject"
                    'styl.Font.Size = 10.5
                Case "Table text", "Table side heading 1", "Table side heading 2", "Table list bullet 1", "Table list bullet 2", "Table list bullet 3"
                    'styl.Font.Size = 8.5
                Case "Box Text", "Box Text (Bold Italic)", "Box Side Heading 1", "Box Side Heading 2", "Box List Bullet", "Box List Bullet 2", "Box List Bullet 3", "Box List Number", "Box List Number 2", "Box List Number 3"
                    'styl.Font.Size = 8.5
                Case "Box Quote", "Box Quote List Bullet", "Box Quote Source"
                    'styl.Font.Size = 7.5
                Case "Source", "Note"
                    'styl.ParagraphFormat.SpaceBefore = 6.0
                Case "Caption"
                    'styl.Font.Size = 9
                    'styl.ParagraphFormat.SpaceBefore = 12.0
                Case "Cp Report Date", "Cp Report To", "Cp Client Name"
                    'styl.Font.Size = 12
                    'styl.Font.Color = RGB(0, 0, 0)
                    'styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                Case "Cp Title"
                    'styl.Font.Size = 36
                    styl.Font.Color = RGB(0, 0, 0)
                    'styl.Font.Bold = True
                    'styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                    'styl.ParagraphFormat.SpaceBefore = 12.0
                   ' styl.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1
                Case "Cp SubTitle"
                    'styl.Font.Size = 20
                    styl.Font.Color = RGB(0, 0, 0)
                    'styl.Font.Bold = False
                    'styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                Case "Cp About Acil Allen"
                    'styl.Font.Size = 14
                    'styl.Font.Size = 12
                    'styl.Font.Color = RGB(0, 0, 0)
                    styl.Font.Bold = False
                    'styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                Case "Cp Disclaimer", "Cp Disclaimer 9pt"
                    'styl.Font.Size = 10
                    'styl.Font.Color = RGB(0, 0, 0)
                    styl.Font.Bold = False
                    'styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle

                Case "Header-Company Name", "Header"
                    styl.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
                Case "Part - Number"
                    'styl.Font.Color = RGB(0, 0, 0)

                'Case "Part xx"
                    'styl.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1
                    'Make sure the small text in the Part Header is set to the Back colour
                    'styl.Font.Color = Me._wcag_backcolour
                Case "Part - Heading (Banner)", "App - Divider (Heading)"
                    'styl.Font.Color = RGB(0, 0, 0)
                    styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                    styl.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1
                Case "Part - SubHead (Banner)", "App - Divider (Sub Heading)"
                    'styl.Font.Color = RGB(0, 0, 0)
                    styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle

                Case "Heading (Chapter)", "Heading (Appendix)"
                    styl.ParagraphFormat.LineSpacing = 90
                    styl.Font.Position = -6
                Case "SideNote (Italic Left)", "SideNote (Italic Right)", "SideNote (Regular Left)", "SideNote (Regular Right)"
                    styl.Font.Color = RGB(0, 0, 0)
                    styl.Font.Size = 10
                Case "Introduction"
                    'styl.Font.Color = RGB(0, 0, 0)
                Case "Sub Heading 1" 'Found in Landscape
                    'styl.Font.Color = RGB(0, 0, 0)
                Case "Heading 1", "Heading 1 (ES)", "Heading 1 (AP)", "Heading (glossary)"
                    'styl.Font.Color = RGB(255, 255, 255)
                    styl.Font.Color = RGB(0, 0, 0)
                    styl.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1
                    styl.Font.Size = 13
                    '
                Case "Heading 2", "Heading 2 (AP)"
                    styl.Font.Color = Me._wcag_color_heading_purple
                    'styl.Font.Size = 18
                    styl.Font.Size = 14
                    'styl.Font.Color = RGB(0, 0, 0)'
                Case "Heading 2 (ES)", "Heading 2 (no number)"
                    styl.Font.Color = Me._wcag_color_heading_purple
                    styl.Font.Size = 14
                    styl.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2

                Case "Heading 3", "Heading 3 (ES)", "Heading 3 (AP)", "Heading 3 (no number)"
                    'styl.Font.Size = 14
                    styl.Font.Size = 12
                    'styl.Font.Color = RGB(0, 0, 0)'
                Case "Heading 4", "Heading 4 (ES)", "Heading 4 (AP)", "Heading 4 (no number)"
                    'styl.Font.Size = 14
                    styl.Font.Size = 11
                    'styl.Font.Color = RGB(0, 0, 0)'
                Case "Heading 5", "Heading 5 (ES)", "Heading 5 (AP)"
                    'styl.Font.Size = 13
                    styl.Font.Size = 11
                    'styl.Font.Color = RGB(0, 0, 0)'
                Case "Heading 6", "Heading 6 (ES)", "Heading 6 (AP)"
                    'styl.Font.Size = 13
                    styl.Font.Size = 11
                    'styl.Font.Color = RGB(0, 0, 0)


                'Case "Part - Number"
                    'styl.Font.Size = 80  
                Case "TOC Heading"
                    styl.Font.Color = Me._wcag_color_heading_purple
                    styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                    styl.Font.Size = 26
                    styl.ParagraphFormat.SpaceBefore = 10.0
                    styl.ParagraphFormat.SpaceAfter = 40.0
                    styl.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1
                Case "TOC 1"
                    styl.Font.Color = Me._wcag_color_heading_purple
                    styl.Font.Size = 12
                    styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast
                    styl.ParagraphFormat.LineSpacing = 14
                    styl.ParagraphFormat.SpaceBefore = 12.0
                    styl.ParagraphFormat.SpaceAfter = 6.0
                Case "TOC 2", "TOC TOFSubHeading"
                    styl.Font.Size = 12
                    'styl.Font.Color = RGB(0, 0, 0)
                    styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast
                    styl.ParagraphFormat.LineSpacing = 14
                    styl.ParagraphFormat.SpaceBefore = 8.0
                    styl.ParagraphFormat.SpaceAfter = 4.0
                Case "TOC 3", "TOC General"
                    styl.Font.Size = 11
                    'styl.Font.Color = RGB(0, 0, 0)
                    styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast
                    styl.ParagraphFormat.LineSpacing = 14
                Case "TOC 4"
                    styl.Font.Size = 10
                    'styl.Font.Color = RGB(0, 0, 0)
                    styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast
                    styl.ParagraphFormat.LineSpacing = 13
                Case "Footer Text"
                    styl.Font.Size = 10
                    'styl.Font.Color = RGB(0, 0, 0)
                Case "Table column headings"
                    styl.Font.Color = RGB(255, 255, 255)


            End Select
        Next
        '
        'Now do Table Styles
        Me.wcag_styles_addHeading1AltStyle(myDoc)
        '
        Try
            Me.wcag_stylesTable_addAACBasic(myDoc)
            'Me.wcag_stylesTable_addAACWide(myDoc)  'Not necessary since we'll use the basic table style for aac tables
            Me.wcag_stylesTable_addBanner(myDoc)    'Now placed In 'wcag_convert_bannersToWCAG'
            Me.wcag_stylesTable_addTblHeader(myDoc)
            Me.wcag_stylesTable_addBodyWithLines(myDoc)
            Me.wcag_stylesTable_addAACBox(myDoc)
            Me.wcag_stylesTable_addAACFigure(myDoc)

        Catch ex As Exception
            MsgBox("Error in styles_Table_add")
        End Try
        '
finis:
        '
    End Sub
    '
    '
    '
    ''' <summary>
    ''' This method will convert the selected document to comply (partially) with
    ''' WCAG requirements
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub wcag_styles_setForWCAG_old(ByRef myDoc As Word.Document)
        Dim styl As Word.Style
        Dim objGlobals As New cGlobals()
        '

        styl = myDoc.Styles.Item("Normal")
        styl.Font.Size = 9.5
        'styl.Font.Size = 10
        styl.Font.Name = "Arial"
        styl.Font.Color = RGB(0, 0, 0)
        '
        'ByPass the styles conversion if this document has already been through the
        'WCAG conversion
        'If Me.wcag_docProps_isAccessible() Then Exit Sub
        '
        For Each styl In myDoc.Styles
            styl.Font.Color = RGB(0, 0, 0)
            If styl.Font.Name <> "Arial" Then styl.Font.Name = "Arial"
            '
            GoTo finis
            '
            Select Case styl.NameLocal
                Case "Normal"
                    'styl.Font.Size = 11
                    'styl.Font.Color = RGB(0, 0, 0)
                    'styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast
                    'styl.ParagraphFormat.LineSpacing = 16
                Case "LetterSubject"
                    'styl.Font.Size = 10.5
                Case "Table text", "Table side heading 1", "Table side heading 2", "Table list bullet 1", "Table list bullet 2", "Table list bullet 3"
                    'styl.Font.Size = 8.5
                Case "Box Text", "Box Text (Bold Italic)", "Box Side Heading 1", "Box Side Heading 2", "Box List Bullet", "Box List Bullet 2", "Box List Bullet 3", "Box List Number", "Box List Number 2", "Box List Number 3"
                    'styl.Font.Size = 8.5
                Case "Box Quote", "Box Quote List Bullet", "Box Quote Source"
                    'styl.Font.Size = 7.5
                Case "Source", "Note"
                    styl.ParagraphFormat.SpaceBefore = 6.0
                Case "Caption"
                    'styl.Font.Size = 9
                    styl.ParagraphFormat.SpaceBefore = 12.0
                Case "Cp Report Date", "Cp Report To", "Cp Client Name"
                    'styl.Font.Size = 12
                    'styl.Font.Color = RGB(0, 0, 0)
                    styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                Case "Cp Title"
                    'styl.Font.Size = 36
                    styl.Font.Color = RGB(0, 0, 0)
                    'styl.Font.Bold = True
                    'styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                    'styl.ParagraphFormat.SpaceBefore = 12.0
                   ' styl.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1
                Case "Cp SubTitle"
                    'styl.Font.Size = 20
                    styl.Font.Color = RGB(0, 0, 0)
                    'styl.Font.Bold = False
                    'styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                Case "Cp About Acil Allen"
                    'styl.Font.Size = 14
                    'styl.Font.Size = 12
                    'styl.Font.Color = RGB(0, 0, 0)
                    styl.Font.Bold = False
                    'styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                Case "Cp Disclaimer", "Cp Disclaimer 9pt"
                    'styl.Font.Size = 10
                    'styl.Font.Color = RGB(0, 0, 0)
                    styl.Font.Bold = False
                    'styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle

                Case "Header-Company Name", "Header"
                    styl.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
                Case "Part - Number"
                    'styl.Font.Color = RGB(0, 0, 0)

                'Case "Part xx"
                    'styl.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1
                    'Make sure the small text in the Part Header is set to the Back colour
                    'styl.Font.Color = Me._wcag_backcolour
                Case "Part - Heading (Banner)", "App - Divider (Heading)"
                    'styl.Font.Color = RGB(0, 0, 0)
                    styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                    styl.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1
                Case "Part - SubHead (Banner)", "App - Divider (Sub Heading)"
                    'styl.Font.Color = RGB(0, 0, 0)
                    styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle

                Case "Heading (Chapter)", "Heading (Appendix)"
                    styl.ParagraphFormat.LineSpacing = 90
                    styl.Font.Position = -6
                Case "SideNote (Italic Left)", "SideNote (Italic Right)", "SideNote (Regular Left)", "SideNote (Regular Right)"
                    styl.Font.Color = RGB(0, 0, 0)
                    styl.Font.Size = 10
                Case "Introduction"
                    'styl.Font.Color = RGB(0, 0, 0)
                Case "Sub Heading 1" 'Found in Landscape
                    'styl.Font.Color = RGB(0, 0, 0)
                Case "Heading 1", "Heading 1 (ES)", "Heading 1 (AP)", "Heading (glossary)"
                    'styl.Font.Color = RGB(255, 255, 255)
                    styl.Font.Color = RGB(0, 0, 0)
                    styl.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1
                    styl.Font.Size = 13
                    '
                Case "Heading 2", "Heading 2 (AP)"
                    styl.Font.Color = Me._wcag_color_heading_purple
                    'styl.Font.Size = 18
                    styl.Font.Size = 14
                    'styl.Font.Color = RGB(0, 0, 0)'
                Case "Heading 2 (ES)", "Heading 2 (no number)"
                    styl.Font.Color = Me._wcag_color_heading_purple
                    styl.Font.Size = 14
                    styl.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2

                Case "Heading 3", "Heading 3 (ES)", "Heading 3 (AP)", "Heading 3 (no number)"
                    'styl.Font.Size = 14
                    styl.Font.Size = 12
                    'styl.Font.Color = RGB(0, 0, 0)'
                Case "Heading 4", "Heading 4 (ES)", "Heading 4 (AP)", "Heading 4 (no number)"
                    'styl.Font.Size = 14
                    styl.Font.Size = 11
                    'styl.Font.Color = RGB(0, 0, 0)'
                Case "Heading 5", "Heading 5 (ES)", "Heading 5 (AP)"
                    'styl.Font.Size = 13
                    styl.Font.Size = 11
                    'styl.Font.Color = RGB(0, 0, 0)'
                Case "Heading 6", "Heading 6 (ES)", "Heading 6 (AP)"
                    'styl.Font.Size = 13
                    styl.Font.Size = 11
                    'styl.Font.Color = RGB(0, 0, 0)


                'Case "Part - Number"
                    'styl.Font.Size = 80  
                Case "TOC Heading"
                    styl.Font.Color = Me._wcag_color_heading_purple
                    styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                    styl.Font.Size = 26
                    styl.ParagraphFormat.SpaceBefore = 10.0
                    styl.ParagraphFormat.SpaceAfter = 40.0
                    styl.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1
                Case "TOC 1"
                    styl.Font.Color = Me._wcag_color_heading_purple
                    styl.Font.Size = 12
                    styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast
                    styl.ParagraphFormat.LineSpacing = 14
                    styl.ParagraphFormat.SpaceBefore = 12.0
                    styl.ParagraphFormat.SpaceAfter = 6.0
                Case "TOC 2", "TOC TOFSubHeading"
                    styl.Font.Size = 12
                    'styl.Font.Color = RGB(0, 0, 0)
                    styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast
                    styl.ParagraphFormat.LineSpacing = 14
                    styl.ParagraphFormat.SpaceBefore = 8.0
                    styl.ParagraphFormat.SpaceAfter = 4.0
                Case "TOC 3", "TOC General"
                    styl.Font.Size = 11
                    'styl.Font.Color = RGB(0, 0, 0)
                    styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast
                    styl.ParagraphFormat.LineSpacing = 14
                Case "TOC 4"
                    styl.Font.Size = 10
                    'styl.Font.Color = RGB(0, 0, 0)
                    styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast
                    styl.ParagraphFormat.LineSpacing = 13
                Case "Footer Text"
                    styl.Font.Size = 10
                    'styl.Font.Color = RGB(0, 0, 0)
                Case "Table column headings"
                    styl.Font.Color = RGB(255, 255, 255)


            End Select
        Next
        '
        'Now do Table Styles
        Me.wcag_styles_addHeading1AltStyle(myDoc)
        '
        Try
            Me.wcag_stylesTable_addAACBasic(myDoc)
            'Me.wcag_stylesTable_addAACWide(myDoc)  'Not necessary since we'll use the basic table style for aac tables
            Me.wcag_stylesTable_addBanner(myDoc)    'Now placed In 'wcag_convert_bannersToWCAG'
            Me.wcag_stylesTable_addTblHeader(myDoc)
            Me.wcag_stylesTable_addBodyWithLines(myDoc)
            Me.wcag_stylesTable_addAACBox(myDoc)
            Me.wcag_stylesTable_addAACFigure(myDoc)

        Catch ex As Exception
            MsgBox("Error in styles_Table_add")
        End Try
        '
finis:
        '
    End Sub
    '
#Region "Styles"

    Public Sub wcag_styles_addHeading1AltStyle(ByRef myDoc As Word.Document)
        Dim styl As Word.Style
        Dim objGlobals As New cGlobals()
        '
        'If the style does not exist.. add it, then modify
        Try
            styl = myDoc.Styles.Item("Heading 1 (wcag)")
        Catch ex As Exception
            styl = myDoc.Styles.Add("Heading 1 (wcag)", WdStyleType.wdStyleTypeParagraph)
            styl.BaseStyle = myDoc.Styles.Item("Normal - no space")
        End Try
        '
        styl.Font.Color = RGB(255, 254, 255)
        styl.Font.Size = 26
        styl.Font.Bold = False
        styl.ParagraphFormat.LeftIndent = 6.8
        styl.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
        styl.ParagraphFormat.LineSpacing = 26
        styl.ParagraphFormat.SpaceBefore = 0.0
        styl.ParagraphFormat.SpaceAfter = 10.0

    End Sub
    '
    '

    Public Sub wcag_stylesTable_addTblHeader(ByRef myDoc As Word.Document)
        Dim objStylesMgr As New cStylesManager()
        Dim styl As Word.Style
        Dim stylCon As Word.ConditionalStyle
        Dim stylTbl As Word.TableStyle
        Dim objGlobals As New cGlobals()
        '
        'If the style does not exist.. add it, then modify
        styl = objStylesMgr.style_get_style("aac Table (tblHeader)", myDoc)
        If IsNothing(styl) Then
            'MsgBox("No aac Chapter Banner")
            styl = myDoc.Styles.Add("aac Table (tblHeader)", WdStyleType.wdStyleTypeTable)
        Else
            'MsgBox("Found aac Chapter Banner")
        End If
        '
        '
        'styl.BaseStyle = myDoc.Styles.Item("aac Table (no lines)")
        styl.BaseStyle = objStylesMgr.style_get_style("aac Table (no lines)", myDoc)

        styl.Font.Color = RGB(255, 254, 255)
        '
        'Formatting for the table
        stylTbl = styl.Table
        stylTbl.Shading.ForegroundPatternColor = Me._wcag_color_backcolour_tblHeader
        '
        stylTbl.Alignment = WdRowAlignment.wdAlignRowLeft
        stylTbl.TopPadding = 2.0
        stylTbl.BottomPadding = 2.0
        'stylTbl.LeftIndent = -22.7
        'styl
        stylCon = styl.Table.Condition(WdConditionCode.wdFirstRow)
        'stylCon = styl.Table.Condition(WdConditionCode.w)

        'stylTable.Table.
        'stylTbl.rcells = WdCellVerticalAlignment.wdCellAlignVerticalBottom

    End Sub
    '
    Public Sub wcag_stylesTable_addBodyWithLines(ByRef myDoc As Word.Document)
        Dim objStylesMgr As New cStylesManager()
        Dim styl As Word.Style
        Dim stylTbl As Word.TableStyle
        Dim brdr As Word.Border
        Dim stylCond As Word.ConditionalStyle
        Dim objGlobals As New cGlobals()
        '
        'If the style does not exist.. add it, then modify
        styl = objStylesMgr.style_get_style("aac Table (with lines)", myDoc)
        If IsNothing(styl) Then
            'MsgBox("No aac Chapter Banner")
            styl = myDoc.Styles.Add("aac Table (with lines)", WdStyleType.wdStyleTypeTable)
        Else
            'MsgBox("Found aac Chapter Banner")
        End If
        '
        '
        'styl.BaseStyle = myDoc.Styles.Item("aac Table (no lines)")
        styl.BaseStyle = objStylesMgr.style_get_style("aac Table (no lines)", myDoc)
        styl.Font.Name = "Arial"
        '
        styl.Font.Color = RGB(0, 0, 0)
        'styl.Table.Shading.ForegroundPatternColor = RGB(54, 31, 76)
        '
        styl.Table.Alignment = WdRowAlignment.wdAlignRowLeft
        '
        stylTbl = styl.Table
        stylTbl.BottomPadding = 2.0
        '
        'brdr = stylTbl.Borders.Item(WdBorderType.wdBorderTop)
        'brdr.LineStyle = WdLineStyle.wdLineStyleSingle
        'brdr.LineWidth = WdLineWidth.wdLineWidth050pt
        'brdr.Color = RGB(255, 0, 0)
        '
        brdr = stylTbl.Borders.Item(WdBorderType.wdBorderHorizontal)
        brdr.LineStyle = WdLineStyle.wdLineStyleSingle
        brdr.LineWidth = WdLineWidth.wdLineWidth050pt
        brdr.Color = RGB(0, 0, 0)
        '
        brdr = stylTbl.Borders.Item(WdBorderType.wdBorderBottom)
        brdr.LineStyle = WdLineStyle.wdLineStyleSingle
        brdr.LineWidth = WdLineWidth.wdLineWidth050pt
        brdr.Color = RGB(0, 0, 0)
        '
        stylCond = styl.Table.Condition(WdConditionCode.wdFirstRow)

        '
        'stylCond.
        'stylTable.Table.
        'stylTable.Table.cells = WdCellVerticalAlignment.wdCellAlignVerticalBottom

    End Sub
    '
    '
    ''' <summary>
    ''' This method will add both the standard, 'aac Table (Figure)' and wide, 'aac Table (Figure-Wide)' figure table styles
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub wcag_stylesTable_addAACFigure(ByRef myDoc As Word.Document)
        Dim stylTable As Word.Style
        Dim objGlobals As New cGlobals()
        Dim brdrs As Word.Borders
        '       
        'If the style does not exist.. add it, then modify
        Try
            stylTable = myDoc.Styles.Item("aac Table (Figure)")
        Catch ex As Exception
            stylTable = myDoc.Styles.Add("aac Table (Figure)", WdStyleType.wdStyleTypeTable)
        End Try
        '
        stylTable.BaseStyle = myDoc.Styles.Item("aac Table (no lines)")
        stylTable.Table.TopPadding = 0
        stylTable.Table.BottomPadding = 0
        stylTable.Table.LeftPadding = 0
        stylTable.Table.RightPadding = 0
        '
        brdrs = stylTable.Table.Borders
        'brdrs.InsideLineStyle = WdLineStyle.wdLineStyleNone
        'brdrs.Item(WdBorderType.wdBorderTop).LineStyle = WdLineStyle.wdLineStyleSingle
        'brdrs.Item(WdBorderType.wdBorderTop).LineWidth = WdLineWidth.wdLineWidth050pt
        'brdrs.Item(WdBorderType.wdBorderTop).Color = RGB(0, 0, 0)
        '
        'brdrs.Item(WdBorderType.wdBorderHorizontal).LineStyle = WdLineStyle.wdLineStyleSingle
        'brdrs.Item(WdBorderType.wdBorderHorizontal).LineWidth = WdLineWidth.wdLineWidth050pt
        'brdrs.Item(WdBorderType.wdBorderHorizontal).Color = RGB(0, 0, 0)
        '
        'brdrs.Item(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleSingle
        'brdrs.Item(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth050pt
        'brdrs.Item(WdBorderType.wdBorderBottom).Color = RGB(0, 0, 0)
        '
        'If the style does not exist.. add it, then modify
        Try
            stylTable = myDoc.Styles.Item("aac Table (Figure-Wide)")
        Catch ex As Exception
            stylTable = myDoc.Styles.Add("aac Table (Figure-Wide)", WdStyleType.wdStyleTypeTable)
        End Try
        '
        stylTable.BaseStyle = myDoc.Styles.Item("aac Table (Figure)")
        stylTable.Table.LeftIndent = -98.15
        'tblTop.Style.Table.LeftIndent = leftIndent
        '
finis:
    End Sub


    '

    Public Sub wcag_stylesTable_addBanner(ByRef myDoc As Word.Document)
        Dim objStylesMgr As New cStylesManager()
        Dim styl As Word.Style
        Dim objGlobals As New cGlobals()
        '
        'If the style does not exist.. add it, then modify
        styl = objStylesMgr.style_get_style("aac Chapter Banner", myDoc)
        If IsNothing(styl) Then
            'MsgBox("No aac Chapter Banner")
            styl = myDoc.Styles.Add("aac Chapter Banner", WdStyleType.wdStyleTypeTable)
        Else
            'MsgBox("Found aac Chapter Banner")
        End If
        '
        '
        'styl.BaseStyle = myDoc.Styles.Item("aac Table (no lines)")
        styl.BaseStyle = objStylesMgr.style_get_style("aac Table (no lines)", myDoc)
        styl.Font.Name = "Arial"
        '
        styl.Font.Color = RGB(255, 254, 255)
        styl.Table.Shading.ForegroundPatternColor = RGB(54, 31, 76)
        '
        styl.Table.Alignment = WdRowAlignment.wdAlignRowLeft
        styl.Table.Condition(WdConditionCode.wdFirstRow)
        'stylTable.Table.
        'stylTable.Table.cells = WdCellVerticalAlignment.wdCellAlignVerticalBottom

    End Sub
    '
    Public Sub wcag_stylesTable_addAACBasic(ByRef myDoc As Word.Document)
        Dim styl As Word.Style
        'Dim stylTable As Word.TableStyle
        Dim objGlobals As New cGlobals()
        '
        Try
            styl = myDoc.Styles.Item("aac Table (Basic)")
        Catch ex As Exception
            styl = myDoc.Styles.Add("aac Table (Basic)", WdStyleType.wdStyleTypeTable)
            styl.BaseStyle = myDoc.Styles.Item("aac Table (no lines)")
        End Try
        '
        styl.Font.Name = "Arial"
        'stylTable = styl
        '
        'stylTable.BaseStyle = myDoc.Styles.Item("Table Normal")
        'stylCond = stylTable.Condition(WdConditionCode.wdFirstRow)

    End Sub

    '
    Public Sub wcag_stylesTable_addAACWide(ByRef myDoc As Word.Document)
        Dim styl As Word.Style
        'Dim stylTable As Word.TableStyle
        Dim objGlobals As New cGlobals()
        '
        Try
            styl = myDoc.Styles.Item("aac Table (Wide)")
        Catch ex As Exception
            styl = myDoc.Styles.Add("aac Table (Wide)", WdStyleType.wdStyleTypeTable)
        End Try
        '
        styl.BaseStyle = myDoc.Styles.Item("aac Table (Basic)")
        styl.Table.LeftIndent = Me._wcag_tables_leftIndent_Wide
        '
    End Sub
    '
    ''' <summary>
    ''' This one works with the original code whereas the prior add Table Styles seemed to
    ''' fault on myDoc.Styles.Item or myDoc.Styles.Add
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub wcag_stylesTable_addAACBox(ByRef myDoc As Word.Document)
        Dim stylTable As Word.Style
        Dim objGlobals As New cGlobals()
        '
        'If the style does not exist.. add it, then modify
        Try
            stylTable = myDoc.Styles.Item("aac Table (Box)")
        Catch ex As Exception
            stylTable = myDoc.Styles.Add("aac Table (Box)", WdStyleType.wdStyleTypeTable)
            stylTable.BaseStyle = myDoc.Styles.Item("aac Table (no lines)")
        End Try
        '
        stylTable.Table.Shading.ForegroundPatternColor = RGB(229, 229, 229)
        stylTable.Table.TopPadding = 5.6
        stylTable.Table.BottomPadding = 9.2
        stylTable.Table.LeftPadding = 0
        stylTable.Table.RightPadding = 6
        'tblTop.Style.Table.LeftIndent = leftIndent
        '
    End Sub
    '
#End Region
    '
    Public Sub wcag_convert_backColour(ByRef myDoc As Word.Document)
        Dim objGlobals As New cGlobals()
        Dim objWCAGMgr As New cWCAGMgr()
        Dim objCpMgr As New cCoverPageMgr()
        Dim objContsMgr As New cContactsMgr()
        Dim objTOCMgr As New cTOCMgr()
        Dim objDivMgr As New cChptDivider()
        Dim objLogosMgr As New cLogosMgr()
        Dim objhfMgr As New cHeaderFooterMgr()
        Dim objPnlMgr As New cBackPanelMgr()
        Dim lstTblEdges As New Collection()
        Dim strDivType As String
        Dim tbl As Word.Table
        Dim tblWidth As Single
        'Dim myDoc As Word.Document
        Dim sect As Word.Section
        Dim hf As Word.HeaderFooter
        Dim shp As Word.Shape
        'Dim inlineShp As Word.InlineShape
        Dim fillColour As Long

        'https://stackoverflow.com/questions/53042897/set-shape-inlineshape-as-decorative-in-word

        'myDoc = objGlobals.glb_get_wrdActiveDoc()
        tblWidth = 0.0
        strDivType = ""
        tbl = Nothing
        shp = Nothing

        For Each sect In myDoc.Sections
            If objCpMgr.cp_Bool_IsCoverPage(sect) Then
                objLogosMgr.logos_set_colour(sect, RGB(0, 0, 0), -1)
            End If
            '
            If objContsMgr.is_ContactsPage_Front(sect) Then
                Me.wcag_contactsPageFront_removeTables(myDoc)
                'sect.PageSetup.TopMargin = sect.PageSetup.TopMargin + 60.0
            End If
            '
            If objContsMgr.is_ContactsPage_Back(sect) Then
                Me.wcag_contactsPageBack_removeTables(sect)
                'sect.PageSetup.TopMargin = sect.PageSetup.TopMargin + 60.0
            End If
            '
            '
            If objDivMgr.is_divider_Any(sect) Then
                'hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                'objLogosMgr.logos_set_colour(hf, RGB(0, 0, 0), -1)
            End If




            For Each hf In sect.Headers
                If hf.Exists Then
                    For Each shp In hf.Shapes
                        Select Case shp.Name
                            Case objPnlMgr.strShapeName, objPnlMgr.strBackPanel_CaseStudy
                                'fillColour = shp.Fill.ForeColor.RGB
                                'Me.wcag_backColour_BorderAndFill(shp)
                                objPnlMgr.pnl_reset_BackPanelTransparency(0.25, sect)
                                'fillColour = RGB(20, 0, 52)
                                'shp.Line.Visible = True
                                'shp.Line.Style = Microsoft.Office.Core.MsoLineStyle.msoLineSingle
                                'shp.Line.ForeColor.RGB = fillColour
                                'shp.Line.Weight = 25
                                '
                                'shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
                                'shp.Fill.ForeColor.RGB = RGB(196, 172, 221)
                                'shp.Fill.ForeColor.RGB = Me._wcag_backcolour
                                '
                                'Set the back ccolour Decorative property
                                'objWCAGMgr.wcag_set_decorative(shp, True)
                                '
                                'inlineShp = shp.ConvertToInlineShape()
                            Case "aac_jigsaw_Wide"
                                fillColour = RGB(20, 0, 52)
                                shp.Fill.ForeColor.RGB = fillColour

                        End Select
                    Next
                End If
                'shp = hf.Range.ShapeRange().Item("aac_BackColour")
                'shp = hf.Shapes.Item("aac_BackColour")
                'If Not IsNothing(shp) Then
                'MsgBox(shp.Name)
                'fillColour = shp.Fill.ForeColor.RGB
                'shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
                'shp.Line.Style = Microsoft.Office.Core.MsoLineStyle.msoLineSingle
                'shp.Line.ForeColor.RGB = fillColour
                'shp.Line.Weight = 50
                'End If
            Next
        Next

    End Sub

    '
    Public Sub x_wcag_convert_backColour(ByRef myDoc As Word.Document)
        Dim objGlobals As New cGlobals()
        Dim objWCAGMgr As New cWCAGMgr()
        Dim objCpMgr As New cCoverPageMgr()
        Dim objContsMgr As New cContactsMgr()
        Dim objTOCMgr As New cTOCMgr()
        Dim objDivMgr As New cChptDivider()
        Dim objLogosMgr As New cLogosMgr()
        Dim objhfMgr As New cHeaderFooterMgr()
        Dim lstTblEdges As New Collection()
        Dim strDivType As String
        Dim tbl As Word.Table
        Dim tblWidth As Single
        'Dim myDoc As Word.Document
        Dim sect As Word.Section
        Dim hf As Word.HeaderFooter
        Dim shp As Word.Shape
        'Dim inlineShp As Word.InlineShape
        Dim fillColour As Long

        'https://stackoverflow.com/questions/53042897/set-shape-inlineshape-as-decorative-in-word

        'myDoc = objGlobals.glb_get_wrdActiveDoc()
        tblWidth = 0.0
        strDivType = ""
        tbl = Nothing
        shp = Nothing

        For Each sect In myDoc.Sections
            If objCpMgr.cp_Bool_IsCoverPage(sect) Then
                Me.wcag_cp_removeTable(myDoc)
                '
                objLogosMgr.logos_set_colour(sect, RGB(0, 0, 0), -1)

                'If Not Me.wcag_docProps_isAccessible() Then
                'objCpMgr.do_HeaderLogo_Colour_noHeaderTable(sect, RGB(255, 255, 255), -1)
                'Else
                'objCpMgr.do_HeaderLogo_Colour_noHeaderTable(sect, RGB(0, 0, 0), -1)
                'End If
                '
                hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                If objWCAGMgr.wcag_get_shpInHf(hf, "cp_pict_large", shp) Then
                    objWCAGMgr.wcag_set_decorative(shp, True)
                End If
                If objWCAGMgr.wcag_get_shpInHf(hf, "cp_Empty_Pattern_Small", shp) Then
                    objWCAGMgr.wcag_set_decorative(shp, True)
                End If
                If objWCAGMgr.wcag_get_shpInHf(hf, "cp_pict_large", shp) Then
                    objWCAGMgr.wcag_set_decorative(shp, True)
                End If


            End If
            '
            If objContsMgr.is_ContactsPage_Front(sect) Then
                Me.wcag_contactsPageFront_removeTables(myDoc)
                sect.PageSetup.TopMargin = sect.PageSetup.TopMargin + 60.0
            End If
            '
            If objContsMgr.is_ContactsPage_Back(sect) Then
                Me.wcag_contactsPageBack_removeTables(sect)
                sect.PageSetup.TopMargin = sect.PageSetup.TopMargin + 60.0
            End If
            '
            '
            If objDivMgr.is_divider_Any(sect) Then
                hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                objLogosMgr.logos_set_colour(hf, RGB(0, 0, 0), -1)
            End If




            For Each hf In sect.Headers
                If hf.Exists Then
                    For Each shp In hf.Shapes
                        Select Case shp.Name
                            Case "aac_BackColour"
                                'fillColour = shp.Fill.ForeColor.RGB
                                Me.wcag_backColour_BorderAndFill(shp)
                                'fillColour = RGB(20, 0, 52)
                                'shp.Line.Visible = True
                                'shp.Line.Style = Microsoft.Office.Core.MsoLineStyle.msoLineSingle
                                'shp.Line.ForeColor.RGB = fillColour
                                'shp.Line.Weight = 25
                                '
                                'shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
                                'shp.Fill.ForeColor.RGB = RGB(196, 172, 221)
                                'shp.Fill.ForeColor.RGB = Me._wcag_backcolour
                                '
                                'Set the back ccolour Decorative property
                                'objWCAGMgr.wcag_set_decorative(shp, True)
                                '
                                'inlineShp = shp.ConvertToInlineShape()
                            Case "aac_jigsaw_Wide"
                                fillColour = RGB(20, 0, 52)
                                shp.Fill.ForeColor.RGB = fillColour

                        End Select
                    Next
                End If
                'shp = hf.Range.ShapeRange().Item("aac_BackColour")
                'shp = hf.Shapes.Item("aac_BackColour")
                'If Not IsNothing(shp) Then
                'MsgBox(shp.Name)
                'fillColour = shp.Fill.ForeColor.RGB
                'shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
                'shp.Line.Style = Microsoft.Office.Core.MsoLineStyle.msoLineSingle
                'shp.Line.ForeColor.RGB = fillColour
                'shp.Line.Weight = 50
                'End If
            Next
        Next

    End Sub
    '
    Public Sub wcag_backColour_BorderAndFill(ByRef shp As Word.Shape)
        'Dim borderColour As Long
        Dim objPnlMgr As New cBackPanelMgr()
        '
        Select Case shp.Name
            Case objPnlMgr.strShapeName
                '
                'borderColour = RGB(20, 0, 52)
                'shp.Line.Visible = True
                'shp.Line.Style = Microsoft.Office.Core.MsoLineStyle.msoLineSingle
                'shp.Line.ForeColor.RGB = borderColour
                'shp.Line.Weight = 25
                '
                shp.Fill.Solid()
                shp.Fill.ForeColor.RGB = Me._wcag_backcolour
                '
            Case objPnlMgr.strBackPanel_CaseStudy

        End Select
        '
        Me.wcag_set_decorative(shp, True)
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will return true if it finds a Shape with the name 'strShapeName' in the HeaderFooter. If
    ''' it does return true, then it will reurn that Shape as the referenced variable shp... Some valid names
    ''' are "logo_AAC_TandG", "aac_Cpg_Logo", "aac_BackColour"
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="strShpName"></param>
    ''' <param name="shp"></param>
    ''' <returns></returns>
    Public Function wcag_get_shpInHf(ByRef hf As Word.HeaderFooter, strShpName As String, ByRef shp As Word.Shape) As Boolean
        Dim rslt As Boolean
        Dim rng As Word.Range
        '
        rslt = False
        rng = hf.Range
        For Each shp In rng.ShapeRange
            If shp.Name = strShpName Then
                rslt = True
                Exit For
            End If
        Next
        '
        Return rslt

    End Function
    '
    ''' <summary>
    ''' This method will set the 'rptAccessible' property of myDoc to objPropsMgr.strTrue or
    ''' objPropsMgr.strFalse depending on whether isAccessible is true or false... Tested OK 20220404
    ''' </summary>
    ''' <param name="isAccessible"></param>
    ''' <param name="myDoc"></param>
    Public Sub wcag_docProps_setAccessibity(isAccessible As Boolean, ByRef myDoc As Word.Document)
        Dim objPropsMgr As New cPropertyMgr()
        Dim strSetting As String
        '
        strSetting = objPropsMgr.strFalse
        If isAccessible Then strSetting = objPropsMgr.strTrue
        '
        objPropsMgr.prps_CustomProperty_set(strSetting, "rptAccessible", myDoc)
        '
    End Sub
    '
    ''' <summary>
    ''' This method will return true of false depending on whether the 'rptAccessible' property
    ''' in myDoc is objPropsMgr.strTrue or objPropsMgr.strFalse. If the property does not exist
    ''' it will be created and set to strFalse
    ''' </summary>
    ''' <returns></returns>
    Public Function wcag_docProps_isAccessible() As Boolean
        Dim objPropsMgr As New cPropertyMgr()
        Dim strRslt As String
        Dim rslt As Boolean
        '
        strRslt = objPropsMgr.prps_CustomProperty_get("rptAccessible", objPropsMgr.strFalse)
        rslt = False
        If strRslt = objPropsMgr.strTrue Then rslt = True
        '
        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' This method will delete the ribbon references '_AssemblyName' and '_AssemblyLocation"
    ''' in the custom document properties of myDoc... Tested OK 20220404
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub wcag_rbn_del(ByRef myDoc As Word.Document)
        Dim objPropsMgr As New cPropertyMgr()
        '
        objPropsMgr.prps_rbn_del(myDoc)
        '
    End Sub
    '
    ''' <summary>
    ''' This method will set the Assembly references for the AAC production version... Tested OK 20220404
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub wcag_rbn_addAAC(ByRef myDoc As Word.Document, strRbnType As String)
        Dim objPropsMgr As New cPropertyMgr()
        '
        objPropsMgr.prps_rbn_setReferences(strRbnType, myDoc)
    End Sub
    '
    ''' <summary>
    ''' This method will set the Assembly references for the AAC production version... Tested OK 20220404
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub wcag_rbn_addTestPlatform(ByRef myDoc As Word.Document)
        Dim objPropsMgr As New cPropertyMgr()
        '
        objPropsMgr.prps_rbn_setReferences("testMachine", myDoc)
        '
    End Sub
    '
    Public Sub wcag_toc_unlinkFields(ByRef myDoc As Word.Document)
        Dim objFldsMgr As New cFieldsMgr()
        '
        objFldsMgr.flds_tocs_unlink(myDoc)
    End Sub
    '
    Public Sub wcag_convert_headersToText(ByRef myDoc As Word.Document, doAsTable As Boolean)
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim sect As Word.Section

        Try
            '
            For Each sect In myDoc.Sections
                objHFMgr.hf_hfs_convertToWCAGCompliance(sect, doAsTable, "header")
            Next
        Catch ex As Exception

        End Try
    End Sub

    '
    '
    Public Sub wcag_contactsPageBack_removeTables(ByRef sect As Word.Section)
        Dim tbl, tblNested As Word.Table
        Dim objTools As New cTools()
        Dim objLegal As New cLegalAndAbout()
        Dim rng As Word.Range
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        Dim para As Word.Paragraph
        '
        Try
            'tbl = sect.Range.Tables.Item(1)
            'tbl.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
            '
            tbl = sect.Range.Tables.Item(1)
            'tbl.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)


            dr = tbl.Rows.Item(2)
            tblNested = dr.Cells.Item(1).Tables.Item(1)
            tblNested.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
            '
            'rng = sect.Range
            'rng.Font.Color = RGB(0, 0, 0)
            'drCellHost = tbl.Range.Cells.Item(2)
            'tblNest = drCellHost.Range.Tables.Item(1)
            'tblNest.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
            '

            '
            drCell = sect.Range.Tables.Item(1).Range.Cells.Item(2)
            rng = drCell.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            'rng.Move(WdUnits.wdCharacter, -1)
            '
            'rng.MoveStart(WdUnits.wdCharacter, -3)
            rng.Delete()
            '
            rng = drCell.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '

            '
            'rng = tblNest.Range.Cells.Item(4).Range
            objLegal.insert_Back_MelbourneAndCanberra(rng, 1)
            para = drCell.Range.Paragraphs.Last
            rng = para.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'rng = drCell.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)

            'rng = tblNest.Range.Cells.Item(5).Range
            objLegal.insert_Back_SydneyAndPerth(rng, 1)
            para = drCell.Range.Paragraphs.Last
            rng = para.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            objLegal.insert_Back_BrisbaneAndAdelaide(rng, 1)
            para = drCell.Range.Paragraphs.Last
            rng = para.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            objLegal.insert_Back_CompanyAndABN(rng, 1)
            para = drCell.Range.Paragraphs.Last
            rng = para.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            objLegal.insert_Back_WebAddress(rng)



            'rng.Font.Color = RGB(255, 0, 0)
            '
            'strText = drCell.Range.Text
            'paras = drCell.Range.Paragraphs
            'para = paras.Last
            'para = para.Previous
            'para.Range.Delete()
            '
            'para = paras.Last
            'para = para.Previous
            'para.Range.Delete()
            '
            'rng = drCell.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            'rng.Move(WdUnits.wdCharacter, -2)
            '
            'rng.MoveStart(WdUnits.wdCharacter, -1)
            'rng.Delete()
            'rng = para.Range
            'rng.MoveStart(WdUnits.wdParagraph, -1)
            'rng.Delete()
            '
            'objTools.tools_cell_GetText(drCell, True)
            'para = para.Previous
            'para = para.Previous
            '

            '
            'tbl = sect.Range.Tables.Item(1)
            'tbl.Range.Font.Color = WdColor.wdColorRed
            'rng = para.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            'rng.mo
            'rng.MoveStart(WdUnits.wdParagraph, -2)
            'rng.Delete()
            'objGlobals.glb_screen_updateLeaveAsItWas()
            '
            rng = sect.Range
            rng.Font.Color = RGB(0, 0, 0)

        Catch ex As Exception
            'MsgBox("Error in Front Contacts Page conversion")
        End Try
    End Sub

    Public Sub wcag_contactsPageFront_removeTables(ByRef myDoc As Word.Document)
        Dim objContsMgr As New cContactsMgr()
        Dim rng As Word.Range
        Dim dr As Word.Row
        Dim tbl, tblNested As Word.Table
        Dim sect As Word.Section

        '
        sect = myDoc.Sections.Item(1)
        '
        If objContsMgr.conts_get_getContactsPageFront(myDoc, sect) Then
            Try
                tbl = sect.Range.Tables.Item(1)
                'tbl.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)

                dr = tbl.Rows.Item(2)
                tblNested = dr.Cells.Item(1).Tables.Item(1)
                tblNested.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
                '
                dr = tbl.Rows.Item(3)
                tblNested = dr.Cells.Item(1).Tables.Item(1)
                tblNested.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
                '
                rng = sect.Range
                rng.Font.Color = RGB(0, 0, 0)
                'drCellHost = tbl.Range.Cells.Item(2)
                'tblNest = drCellHost.Range.Tables.Item(1)
                'tblNest.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
            Catch ex As Exception
                'MsgBox("Error in Front Contacts Page conversion")
            End Try
        End If
    End Sub
    '
    '
    Public Sub wcag_convert_bannersToWCAG(ByRef myDoc As Word.Document)
        Dim drCell As Word.Cell
        Dim dr As Word.Row
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim objGlobals As New cGlobals()
        Dim objTools As New cTools()
        Dim objTagsMgr As New cTagsMgr()
        'Dim myDoc As Word.Document
        Dim strStyleName As String

        'myDoc = objGlobals.glb_get_wrdDoc()

        Try
            '
            For Each tbl In myDoc.Range.Tables
                strStyleName = objTagsMgr.tags_get_tagStyleName(tbl)
                Select Case strStyleName
                    Case "tag_execBanner", "tag_chapterBanner", "tag_glossary_Chpt", "tag_biblio_Chpt", "tag_refs_Chpt", "tag_appendixChapter"
                        drCell = tbl.Range.Cells.Item(1)
                        rng = drCell.Range
                        rng.Delete()
                        '
                        dr = tbl.Rows.Item(1)
                        dr.Shading.ForegroundPatternColor = RGB(54, 31, 76)
                        dr = tbl.Rows.Item(2)
                        dr.Shading.ForegroundPatternColor = RGB(54, 31, 76)
                        '
                        Try
                            dr = tbl.Rows.Item(3)
                            dr.Delete()
                        Catch ex As Exception

                        End Try

                    Case Else

                End Select
                '
            Next
        Catch ex As Exception

        End Try
    End Sub
    '
    ''' <summary>
    ''' This method will remove the Cover Page Tables title block and any
    ''' controls embedded on the Cover Page
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub wcag_cp_removeTable(ByRef myDoc As Word.Document)
        Dim objCpMgr As New cCoverPageMgr()
        Dim sect As Word.Section
        Dim tbl As Word.Table
        Dim rng As Word.Range
        Dim ctrls As Word.ContentControls
        Dim ctrl As Word.ContentControl

        sect = myDoc.Sections.Item(1)

        If objCpMgr.cp_Bool_HasCoverPage(myDoc, sect) Then
            If sect.Range.Tables.Count > 0 Then
                tbl = sect.Range.Tables.Item(1)
                tbl.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
                rng = sect.Range
                rng.Font.Color = RGB(0, 0, 0)
                '
                ctrls = rng.ContentControls
                For Each ctrl In ctrls
                    ctrl.Delete()
                Next
            End If
        End If

    End Sub
    '
    ''' <summary>
    ''' This method is the entry call for a conversion of myDoc to an 'Accessible' version. The user can select whether
    ''' to have the tables outdented or not
    ''' </summary>
    ''' <param name="doTablesAsOutdented"></param>
    ''' <param name="myDoc"></param>
    Public Sub wcag_doc_ToWCAG_entry(doTablesAsOutdented As Boolean, myDoc As Word.Document)
        'Dim myDoc As Word.Document
        Dim objMsgMgr As New cMessageManager()
        Dim objGlobals As New cGlobals()
        Dim sel As Word.Selection
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim Interval As TimeSpan
        Dim endTime, startTime As Date
        Dim strElapsedTime As String
        '
        Try
            objGlobals.glb_cursors_setToWait()
            startTime = TimeOfDay()
            'myDoc = objGlobals.glb_get_wrdActiveDoc()
            '
            sel = objGlobals.glb_get_wrdSel
            rng = sel.Range
            tbl = objGlobals.glb_get_wrdSelTbl
            If IsNothing(tbl) Then
                Me.wcag_doc_ToWCAG_Worker(doTablesAsOutdented, myDoc)
                'Go back to where you started
                rng.Select()
                '
            Else
                Me.wcag_doc_ToWCAG_Worker(doTablesAsOutdented, myDoc)
                'Go to the first page
                rng = myDoc.Sections.First.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Select()
                '
            End If
            '
            endTime = TimeOfDay()
            Interval = endTime - startTime
            strElapsedTime = Int(Interval.TotalSeconds()) & " Seconds"

            objGlobals.glb_cursors_setToNormal()
            objGlobals.glb_screen_update(True)
            MsgBox("Conversion complete in " + strElapsedTime)
        Catch ex As Exception
            MsgBox("Conversion has failed... Check your document against the 'Styles Guide'")
        End Try

        '
    End Sub
    '
    '
    ''' <summary>
    ''' This is the main worker method for converting myDoc to an 'Accessible' version
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub wcag_doc_ToWCAG_Worker(ByRef myDoc As Word.Document)
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objFldsMgr As New cFieldsMgr()
        Dim objGlobals As New cGlobals()
        Dim objTablesMgr As New cTablesMgr(myDoc)
        Dim objTools As New cTools()
        Dim leftIndent, cellPadding, leftIndentBody As Single
        Dim isFloating, doAsTable, isAccessible As Boolean
        Dim tbl As Word.Table
        Dim dr As Word.Row
        Dim sect As Word.Section
        Dim stylTableOfFigures As Word.Style
        Dim strStyleName, strTableBoxFigureType, strFirstRowText As String

        '
        'objGlobals.glb_cursors_setToWait()
        strTableBoxFigureType = ""
        strFirstRowText = ""
        leftIndent = 0.0
        leftIndentBody = 0.0
        cellPadding = 0.0
        isFloating = False
        doAsTable = False
        '
        Try
            isAccessible = Me.wcag_docProps_isAccessible()
            '
            If Not isAccessible Then
                'Set all styles to black, then set all text to black just to ensure that we have
                'got all hand adjustments
                Call Me.wcag_styles_setForWCAG(myDoc)
                myDoc.Range.Font.Color = RGB(0, 0, 0)
            End If
            '
            sect = Nothing
            stylTableOfFigures = myDoc.Styles.Item("Table of Figures")
            '
            'objGlobals.glb_screen_update()
            myDoc.Fields.Update()
            '
            Me.wcag_convert_backColour(myDoc)
            objGlobals.glb_screen_updateLeaveAsItWas()
            '
            'objFldsMgr.flds_tocs_unlink(myDoc, True)
            '
            '
            'objFldsMgr.flds_tocs_unlink(myDoc, False)
            'objGlobals.glb_set_fieldShading("always")
            '
            'Me.wcag_convert_headersToWCAG(myDoc, doAsTable)
            Me.wcag_convert_footersToWCAG(myDoc, doAsTable)
            'objFldsMgr.flds_footer_unlink(myDoc)
            '
            'If Not isAccessible Then
            'Call Me.wcag_styles_setForWCAG(myDoc)
            'myDoc.Range.Font.Color = RGB(0, 0, 0)
            'End If
            '
            '
            '
            For Each tbl In myDoc.Tables
                'Place an inner Try here becuase the document may have irregular tables
                'in here which the author did not remove/modify

                Try
                    strStyleName = objTools.tools_tbls_getFirstCellStyleName(tbl)
                    strFirstRowText = objTools.tools_cell_GetText(tbl.Range.Cells.Item(1), True)
                    '
                    Select Case strStyleName
                        Case "Table column headings"
                            dr = tbl.Rows.Item(1)
                            objTablesMgr.tbl_colour_set_colourOfRow(dr, objGlobals._glb_colour_CaseStudy_Grey)
                        '
                        Case "Glossary"
                            dr = tbl.Rows.Item(1)
                            dr.Range.Font.Color = RGB(0, 0, 0)
                            objTablesMgr.tbl_colour_set_colourOfRow(dr, objGlobals._glb_colour_CaseStudy_Grey)
                        '
                        Case "Caption", "Caption Label"
                            If strFirstRowText Like "Figure*" Or strFirstRowText Like "Box*" Then
                                dr = tbl.Rows.Item(1)
                                'dr.Range.Font.Color = RGB(0, 0, 0)
                                'objTablesMgr.tbl_colour_set_colourOfRow(dr, RGB(0, 255, 0))

                            End If
                            '
                            If strFirstRowText Like "Table*" Then
                                dr = tbl.Rows.Item(2)
                                dr.Range.Font.Color = RGB(0, 0, 0)
                                'objTablesMgr.tbl_colour_set_colourOfRow(dr, RGB(0, 255, 0))

                            End If

                    End Select
                    '
                Catch ex2 As Exception

                End Try
                '
            Next
            '
            'myDoc.Fields.Unlink()
            '        'Tested OK 20220404
            '
            'Seem to need to do this again
            myDoc.Range.Font.Color = RGB(0, 0, 0)
            Me.wcag_docProps_setAccessibity(True, myDoc)
            '
            'Tested OK 20220404
            'Me.wcag_rbn_del(myDoc)
            '
        Catch ex As Exception
            MsgBox("Conversion Error")
        End Try
        '
        'objGlobals.glb_cursors_setToNormal()
        '
finis:
        '
        objGlobals.glb_screen_update(True)
        '
    End Sub
    '
    '
    '
    ''' <summary>
    ''' This is the main worker method for converting myDoc to an 'Accessible' version
    ''' </summary>
    ''' <param name="doTablesAsOutdented"></param>
    ''' <param name="myDoc"></param>
    Public Sub wcag_doc_ToWCAG_Worker(doTablesAsOutdented As Boolean, ByRef myDoc As Word.Document)
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objFldsMgr As New cFieldsMgr()
        Dim objGlobals As New cGlobals()
        Dim objTablesMgr As New cTablesMgr(myDoc)
        Dim objTools As New cTools()
        Dim numHeaderRows As Integer
        Dim leftIndent, cellPadding, leftIndentBody, bodyWidth As Single
        Dim isFloating, doAsTable, isAccessible As Boolean
        Dim shp As Word.Shape
        Dim iShp As Word.InlineShape
        Dim tbl As Word.Table
        Dim drCell As Word.Cell
        Dim sect As Word.Section
        Dim stylTableOfFigures As Word.Style
        Dim strStyleName, strTableBoxFigureType As String

        '
        'objGlobals.glb_cursors_setToWait()
        strTableBoxFigureType = ""
        leftIndent = 0.0
        leftIndentBody = 0.0
        cellPadding = 0.0
        isFloating = False
        doAsTable = True
        '
        Try
            isAccessible = Me.wcag_docProps_isAccessible()

            sect = Nothing
            stylTableOfFigures = myDoc.Styles.Item("Table of Figures")
            '
            'objGlobals.glb_screen_update()
            myDoc.Fields.Update()
            '
            Me.wcag_convert_backColour(myDoc)
            objGlobals.glb_screen_update()
            '
            'objFldsMgr.flds_tocs_unlink(myDoc, True)
            objFldsMgr.flds_tocs_unlink(myDoc, False)
            '
            Me.wcag_convert_headersToWCAG(myDoc, doAsTable)
            Me.wcag_convert_footersToWCAG(myDoc, doAsTable)
            objFldsMgr.flds_footer_unlink(myDoc)
            '
            If Not isAccessible Then Call Me.wcag_styles_setForWCAG(myDoc)
            '
            'Handle the pullouts
            For Each shp In myDoc.Shapes
                If shp.Type = MsoShapeType.msoTextBox Then
                    'if shp.WrapFormat.Type = WdWrapType.wdWrapFront
                    Me.wcag_set_decorative(shp, True)
                    'txtBox = shp
                    If shp.TextFrame.ContainingRange.Tables.Count <> 0 Then
                        For Each tbl In shp.TextFrame.ContainingRange.Tables
                            For Each drCell In tbl.Range.Cells
                                If drCell.Range.InlineShapes.Count <> 0 Then
                                    For Each iShp In drCell.Range.InlineShapes
                                        If iShp.AlternativeText = "" Then
                                            iShp.AlternativeText = "picture in a sidenote"
                                        End If
                                    Next
                                End If
                            Next
                        Next
                    End If
                End If
            Next
            '
            For Each tbl In myDoc.Tables
                strStyleName = objTools.tools_tbls_getFirstCellStyleName(tbl)
                objTablesMgr.tbl_get_tableProperties(tbl, numHeaderRows, leftIndent, cellPadding, leftIndentBody, bodyWidth)
                objTablesMgr.tbl_convert_toInLine(tbl)
                '
                Select Case strStyleName
                    Case "tag_chapterBanner", "tag_execBanner", "tag_glossary_Chpt", "tag_biblio_Chpt", "tag_refs_Chpt", "tag_appendixChapter"
                        Me.wcag_convertbanner_toTableVersion(tbl, strStyleName)
                    Case "tag_partBanner", "tag_appendixPart"
                        tbl.Rows.Item(3).Range.Delete()
                        tbl.Rows.Item(3).Delete()
                        tbl.Rows.Item(1).Delete()
                        tbl.Columns.Item(2).Delete()
                        'tbl.Rows.Item(1).Range.Style = myDoc.Styles.Item("Part - Heading (Banner)")
                        tbl.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
                    Case "PullOut Title", "SideNote*"

                    Case Else
                        drCell = Nothing
                        If objTablesMgr.tbl_is_LegacyAATable(tbl) Then strTableBoxFigureType = "tagTable"
                        If objTablesMgr.tbl_is_AACBox(tbl) Then strTableBoxFigureType = "tagBox"
                        If objTablesMgr.tbl_is_AACFigure(tbl) Then strTableBoxFigureType = "tagFigure"

                        Select Case strTableBoxFigureType
                            Case "tagTable"
                                'wcag_convert_aacTableToWCAG(tbl)
                                wcag_convert_aacTableToWCAG(tbl, doTablesAsOutdented)
                            Case "tagBox"
                                wcag_convert_aacBoxToWCAG(tbl)
                            Case "tagFigure"
                                'wcag_convert_aacFigureToWCAG(tbl)
                                wcag_convert_aacFigureToWCAG_alt(tbl)
                        End Select

                End Select
            Next
            '
            '
bypass:
            '
            'Me.wcag_convert_bannersToWCAG(myDoc)
            'Me.wcag_convert_dividersToWCAG(myDoc)
            'Me.wcag_convert_aacTablesToWCAG(myDoc)
            'Me.wcag_convert_aacBoxesToWCAG(myDoc)
            'Me.wcag_convert_aacFiguresToWCAG(myDoc)
            '
            '
            myDoc.Fields.Unlink()
            '        'Tested OK 20220404
            Me.wcag_docProps_setAccessibity(True, myDoc)
            '
            'Tested OK 20220404
            Me.wcag_rbn_del(myDoc)
            '
        Catch ex As Exception
            MsgBox("Conversion Error")
        End Try
        '
        'objGlobals.glb_cursors_setToNormal()
        '
        '
        objGlobals.glb_screen_update(True)
        '
    End Sub
    '
    '
    Public Sub wcag_convert_headersToWCAG(ByRef myDoc As Word.Document, Optional doAsTable As Boolean = True)
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim sect As Word.Section

        Try
            '
            For Each sect In myDoc.Sections
                objHFMgr.hf_hfs_convertToWCAGCompliance(sect, doAsTable, "header")
            Next
        Catch ex As Exception

        End Try
    End Sub
    '
    '
    Public Sub wcag_convert_footerStylesToWCAG(ByRef myDoc As Word.Document)
        Dim styleFooterText, stylePageNum As Word.Style
        '
        Try
            '
            styleFooterText = myDoc.Styles.Item("Footer Text")
            styleFooterText.Font.Size = 8
            '
            stylePageNum = myDoc.Styles.Item("pageNumber")
            stylePageNum.Font.Size = 14
            '
        Catch ex As Exception

        End Try
        '
    End Sub
    '

    Public Sub wcag_convert_footersToWCAG(ByRef myDoc As Word.Document, Optional doAsTable As Boolean = False)
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim sect As Word.Section
        '
        'First chnage font sizes
        Me.wcag_convert_footerStylesToWCAG(myDoc)
        '
        'The adjust footer by footer to ensure that nay hand colouring is gone
        '
        Try
            '
            For Each sect In myDoc.Sections
                objHFMgr.hf_hfs_convertToWCAGCompliance(sect, doAsTable, "footer")
            Next
            '
        Catch ex As Exception

        End Try
    End Sub
    '
    '
    Public Function wcag_convert_aacFigureToWCAG_alt(ByRef tbl As Word.Table) As Word.Paragraph
        Dim objGlobals As New cGlobals()
        Dim objTablesMgr As New cTablesMgr(tbl.Range.Document)
        Dim objParaMgr As New cParas()
        Dim tblTop, tblBody As Word.Table
        Dim dr, drCustom, drImage As Word.Row
        Dim rng As Word.Range
        Dim paraSource, paraCaption, splitPara As Word.Paragraph
        Dim configuration As Integer
        Dim leftIndent As Single
        '
        rng = Nothing
        paraSource = Nothing
        paraCaption = Nothing
        splitPara = Nothing
        tblTop = Nothing
        tblBody = Nothing
        dr = Nothing
        drImage = Nothing
        drCustom = Nothing
        configuration = 0
        '
        Try
            If objTablesMgr.tbl_is_AACFigure(tbl) Then
                leftIndent = tbl.Rows.LeftIndent
                '
                'Split off the top and bottom rows, when we are finished, tblBody contains the
                'in line picture and any custom commment rows
                '
                If objTablesMgr.tbl_split_Table(2, tbl, splitPara, tblTop) Then
                    rng = tblTop.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
                    paraCaption = rng.Paragraphs.Item(1)
                    objParaMgr.paras_set_HangingIndent(leftIndent, paraCaption)
                    splitPara.Range.Delete()
                    '
                End If
                '
                If objTablesMgr.tbl_find_SourceRow(tbl, dr) Then
                    objTablesMgr.tbl_split_Table(dr, tbl, splitPara, tblBody)
                    rng = tbl.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
                    paraSource = rng.Paragraphs.Item(1)
                    objParaMgr.paras_set_HangingIndent(leftIndent, paraSource)
                    splitPara.Range.Delete()
                End If
                '
                'tblBody now contains the inline picture and any custom comment rows
                '
                Select Case tblBody.Rows.Count
                    Case 1
                        'Standard picture place holder with no custom comment rows
                        objTablesMgr.tbl_apply_figureTableStyle(leftIndent, tblBody)
                        '
                        'Set the top and bottom borders
                        objTablesMgr.tbl_set_borders(tblBody.Rows.First, True, True)
                        '
                        'Now write the Alt text (i.e. if empty we put the Caption text here)
                        dr = objTablesMgr.tbl_find_rowWithInlineShape(tblBody)
                        If Not IsNothing(dr) Then
                            Me.wcag_alttext_write(paraCaption, dr.Range)
                        End If
                    Case Else
                        dr = objTablesMgr.tbl_find_rowWithInlineShape(tblBody)
                        If Not IsNothing(dr) Then
                            configuration = 1                                               'Custom rows at top and bottom
                            If dr.Index = 1 Then configuration = 2                          'Custom rows at bottom only
                            If dr.Index = tblBody.Rows.Last.Index Then configuration = 0    'Custom rows at top only
                            '
                            Select Case configuration
                                Case 0  'Custom rows at top only
                                    'Get the last (or only) custom row at the top and store it's colour
                                    '
                                    drImage = objTablesMgr.tbl_splitTopRow_AACFigureCustomRow(leftIndent, dr, tblBody)
                                    objTablesMgr.tbl_apply_figureTableStyle(leftIndent, tblBody)
                                    objTablesMgr.tbl_set_borders(tblBody.Rows.First, False, True)
                                    '
                                    Me.wcag_alttext_write(paraCaption, drImage.Range)
                                    '
                                Case 1  'Custom rows at top and bottom
                                    'Let's deal with the top first
                                    objTablesMgr.tbl_splitTopRow_AACFigureCustomRow(leftIndent, dr, tblBody)
                                    '
                                    dr = tblBody.Rows.First
                                    drImage = objTablesMgr.tbl_splitBottomRow_AACFigureCustomRow(leftIndent, dr, tblBody)
                                    objTablesMgr.tbl_set_borders(drImage, False, False)
                                    Me.wcag_alttext_write(paraCaption, drImage.Range)

                                    '
                                Case 2  'Custom rows at bottom only
                                    'We have some custom rows
                                    drImage = objTablesMgr.tbl_splitBottomRow_AACFigureCustomRow(leftIndent, dr, tblBody)
                                    Me.wcag_alttext_write(paraCaption, drImage.Range)
                                    '
                            End Select
                            '
                        End If

                End Select


            End If
        Catch ex As Exception

        End Try
        '
finis:
        '
        Return paraSource
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will convert all of the aac outdented Tables in the specified range, rng to wcag
    ''' complaint tables. It successively calls the method 'wcag_convert_aacTableToWCAG(tbl)'.
    ''' The table is checked to see that it is a standard outdented aac table. Typically used to convert
    ''' Tables in a selection
    ''' </summary>
    Public Sub wcag_convert_aacTablesToWCAG(ByRef rng As Word.Range, doTablesAsOutdented As Boolean)
        Dim tbl As Word.Table

        Dim objGlobals As New cGlobals()
        Dim objTools As New cTools()
        '
        Try
            '
            For Each tbl In rng.Tables
                'Me.wcag_convert_aacTableToWCAG(tbl)
                Me.wcag_convert_aacTableToWCAG(tbl, doTablesAsOutdented)
                '
            Next
        Catch ex As Exception

        End Try
    End Sub
    '
    '
    ''' <summary>
    ''' This method allows for quick/easy swapping between approaches
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub wcag_convert_aacTableToWCAG(ByRef tbl As Word.Table, doTablesAsOutdented As Boolean)
        '
        If doTablesAsOutdented Then
            Me.wcag_convert_aacTableToWCAG_Partitioned(tbl)
        Else
            Me.wcag_convert_aacTableToWCAG_noIndent(tbl)

        End If
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method expects as input the Table tbl which is a AAC outdented Table.
    ''' It will convert this to a stgandard regular table and will apply the basic
    ''' Table Style "aac Table (Basic)" with Header Row enabled.. Such Tbales don't
    ''' throw an accessibility error or warning
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub wcag_convert_aacTableToWCAG_noIndent(ByRef tbl As Word.Table)
        Dim objGlobals As New cGlobals()
        Dim objTablesMgr As New cTablesMgr(tbl.Range.Document)
        Dim objParaMgr As New cParas()
        Dim myDoc As Word.Document
        Dim rng As Word.Range
        Dim tblTop As Word.Table
        Dim para, splitPara As Word.Paragraph
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        Dim strCaption As String
        '
        Dim numHeaderRows As Integer
        Dim leftIndent, cellPadding, leftIndentBody, bodyWidth, pageWidth, delta As Single
        '
        myDoc = tbl.Range.Document
        splitPara = Nothing
        para = Nothing
        '
        numHeaderRows = 0
        leftIndent = 0
        cellPadding = 0
        leftIndentBody = 0
        bodyWidth = 0
        '
        pageWidth = 0
        delta = 0
        '
        tblTop = Nothing
        dr = Nothing
        drCell = Nothing
        '
        strCaption = ""

        'Get Table properties..If we are successful then we can operate on it
        If objTablesMgr.tbl_get_tableProperties(tbl, numHeaderRows, leftIndent, cellPadding, leftIndentBody, bodyWidth) Then
            Try
                pageWidth = objGlobals.glb_get_widthBetweenMargins(tbl.Range.Sections.Item(1))
                delta = pageWidth - bodyWidth
                '
                'All WCAG tables must be inline
                objTablesMgr.tbl_convert_toInLine(tbl)
                If numHeaderRows >= 1 Then
                    '
                    'Remove padding and indents... If we apply the basic style at this point with
                    'the indents in place, then entire table is widened to match the first row
                    '
                    For i = 1 To numHeaderRows
                        dr = tbl.Rows.Item(i)
                        drCell = dr.Cells.Item(1)
                        drCell.LeftPadding = 0.0
                        drCell.Width = drCell.Width + leftIndent        'remember leftIndent is negative
                    Next
                    '
                    If bodyWidth <= pageWidth Then
                        For Each dr In tbl.Rows
                            dr.LeftIndent = 0.0
                        Next
                    Else
                        For Each dr In tbl.Rows
                            dr.LeftIndent = delta
                        Next
                    End If
                    '
                    'This forces the table back to a regular structure as defined by the Style
                    tbl.Style = myDoc.Styles.Item("aac Table (Basic)")
                    tbl.ApplyStyleHeadingRows = True
                    '
                    'Re-Colour the Header Rows
                    For i = 1 To numHeaderRows
                        dr = tbl.Rows.Item(i)
                        dr.Shading.ForegroundPatternColor = Me._wcag_color_backcolour_old
                        dr.Shading.BackgroundPatternColor = Me._wcag_color_backcolour_old
                        '
                    Next
                    '
                    'Gte the caption paragraph which is assumed to be directly above the table
                    strCaption = objTablesMgr.tbl_get_tblCaption(tbl, para)
                    '
                    'Now set the indents and adjust the indent of the Caption
                    If bodyWidth <= pageWidth Then
                        For Each dr In tbl.Rows
                            dr.LeftIndent = 0.0
                            objParaMgr.paras_set_HangingIndent(0.0, para)
                        Next
                    Else
                        For Each dr In tbl.Rows
                            dr.LeftIndent = delta
                            objParaMgr.paras_set_HangingIndent(delta, para)
                        Next
                    End If
                    '
                    'Get the Source Rows and do a pre-emptive merge to ensure that we
                    'get consistent behaviour
                    '
                    If objTablesMgr.tbl_find_tableBodyBottom(dr, tbl) Then
                        objTablesMgr.tbl_split_Table(dr, tbl, splitPara, tblTop)
                        For Each dr In tbl.Rows
                            dr.Cells.Merge()
                        Next
                        rng = tbl.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
                        rng.ParagraphFormat.LeftIndent = leftIndentBody
                        splitPara.Range.Delete()
                        '
                    End If
                    '
                    '
                    'Now add left padding for the first column, but do it cell by cell just in case the
                    'author wasn't too careful with alignments
                    For Each dr In tblTop.Rows
                        'dr.Cells.Item(1).LeftPadding = 4
                    Next
                    tbl = tblTop
                End If


            Catch ex As Exception

            End Try
        End If
        '
        'tblIsWide = objTablesMgr.tbl_is_wide(tbl, bodyWidth)
        '
        'Try

        'If objTablesMgr.tbl_is_AACTable(tbl, leftIndent, cellPadding, leftIndentBody) Then

        '
        'For Each dr In tbl.Rows
        'dr.LeftIndent = 0.0
        'Next

        'If tbl.Rows.Item(1).LeftIndent >= 0.0 Then
        'Table has been offset, ofthen happens when converting a floating table to an
        'inline table
        'If Not tblIsWide Then
        'For Each dr In tbl.Rows
        'dr.LeftIndent = dr.LeftIndent - leftIndentBody
        'Next
        'Else
        'For Each dr In tbl.Rows
        'dr.LeftIndent = dr.LeftIndent - (bodyWidth - objGlobals.glb_get_widthBetweenMargins(tbl.Range.Sections.Item(1)) - leftIndent)
        'dr.LeftIndent = dr.LeftIndent - (widthOfTable - objGlobals.glb_get_widthBetweenMargins(tbl.Range.Sections.Item(1)))
        'Next
        'End If
        'Else
        'If tblIsWide Then

        'End If
        'End If
        '
        'GoTo finis
        '
        'If objTablesMgr.tbl_find_tableBodyBottom(dr, tbl) Then
        'objTablesMgr.tbl_split_Table(dr, tbl, splitPara, tblTop)
        'rng = tbl.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
        'rng.ParagraphFormat.LeftIndent = leftIndentBody
        'splitPara.Range.Delete()
        '
        'End If
        '
        'If objTablesMgr.tbl_find_tableBodyTop(dr, tblTop) Then
        'objTablesMgr.tbl_split_Table(dr, tblTop, splitPara, tblHeader)
        '
        'Sometimes we have more than one header row.. Adjust the padding and cell width
        'For Each dr In tblHeader.Rows
        'drCell = dr.Cells.Item(1)
        'drCell.LeftPadding = 0.0
        'drCell.Width = drCell.Width + leftIndent
        'dr.LeftIndent = tblTop.Rows.LeftIndent
        'Next
        '


        '
        'splitPara.Range.Delete()

        '
        'tblHeader.Rows.LeftIndent = 0
        'drCol = tbl.Columns.Item(1)
        'drCol.Width = drCol.Width + leftIndent
        '
        'drCol = tblTop.Columns.First
        'drCol.Width = drCol.Width + leftIndent

        'End If
        '
        'objTablesMgr.tbl_split_Table(2, tblTop, splitPara, tbl
        't 'bl.Range.Cells.Item(1).LeftPadding = 0.0
        '
        'If leftIndent < 0 Then
        'tblTop.Style = myDoc.Styles.Item("aac Table (Basic)")
        '
        'tblTop.Style = myDoc.Styles.Item("aac Table (Wide)")
        'If
        'tblTop.ApplyStyleHeadingRows = True
        '
        'dr = tblTop.Rows.Add(tblTop.Rows.First)
        'dr.Shading.ForegroundPatternColor = objTablesMgr.colourHeader
        'dr.Shading.BackgroundPatternColor = objTablesMgr.colourHeader

        '
        'tbl.Rows.Item(1).Shading.ForegroundPatternColor = objTablesMgr.colourHeader
        'tbl.Rows.Item(1).Shading.BackgroundPatternColor = objTablesMgr.colourHeader
        '
        'dr = tbl.Rows.Item(1)
        'dr.Shading.ForegroundPatternColor = RGB(255, 0, 0)
        'dr.Shading.BackgroundPatternColor = RGB(255, 0, 0)
        '

        'If objTablesMgr.tbl_split_Table(tbl.Rows.Item(1), tbl, splitPara, tblTop) Then

        'End If
        '
        'GoTo finis
        'Make certain all tables are inline
        'If Not objTablesMgr.tbl_get_headerRow(dr, tbl) Then
        'If couldn't find it, we will still proceed assuming that the first row
        ' is the header row
        'dr = tbl.Rows.Item(1)
        'End If
        '
        'If objTablesMgr.tbl_split_Table(2, tbl, splitPara, tblTop) Then
        'GoTo finis
        'drCell = tblTop.Range.Cells.Item(1)
        'leftPadding = drCell.LeftPadding
        'drCell.LeftPadding = 0.0
        'tblTop.Columns.Item(1).Width = tblTop.Columns.Item(1).Width - leftPadding
        'leftIndent = tbl.Rows.LeftIndent
        'tblTop.Rows.LeftIndent = leftIndent
        'splitPara.Range.Delete()
        '
        'GoTo finis
        '
        'dr = objTablesMgr.Get_LastBodyRow_ForSCWC(tbl)
        'If Not dr.Index = 1 Then
        'If objTablesMgr.tbl_split_Table(dr.Index, tbl, splitPara, tblTop) Then
        'rng = tbl.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
        'rng.ParagraphFormat.LeftIndent = leftIndent
        'splitPara.Range.Delete()
        '
        'tblTop.ApplyStyleHeadingRows = True
        'Try
        'If leftIndent < 0 Then
        'tblTop.Style = myDoc.Styles.Item("aac Table (Wide)")
        'Else
        'tblTop.Style = myDoc.Styles.Item("aac Table (Basic)")
        'End If
        '
        '*** Problem, this works but resets the style for all Table
        'tblTop.Style.Table.LeftIndent = leftIndent
        'drCell = tblTop.Range.Cells.Item(1)
        'rng = drCell.Range
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'rng.Select()
        '
        'Catch ex2 As Exception
        'MsgBox("Failed in style application")
        'End Try

        'i = 1
        '
        'End If
        'If
        '
        'Now apply a Table Style. so that we don't get WCAG Read Order error
        '

        'End If
        'End If
        '
        'Catch ex As Exception

        'End Try
finis:
    End Sub
    '
    '
    ''' <summary>
    ''' This method expects as input the Table tbl which is a AAC outdented Table.
    ''' It will split the Table at the Header row and place a 1 pt single line height paragraph
    ''' between the Header and the body. The Source section will be split away and converted to text
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub wcag_convert_aacTableToWCAG_Partitioned(ByRef tbl As Word.Table)
        Dim objGlobals As New cGlobals()
        Dim objTablesMgr As New cTablesMgr(tbl.Range.Document)
        Dim objParaMgr As New cParas()
        Dim myDoc As Word.Document
        Dim rng As Word.Range
        Dim tblTop As Word.Table
        Dim para, splitPara As Word.Paragraph
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        Dim strCaption As String
        '
        Dim numHeaderRows As Integer
        Dim leftIndent, cellPadding, leftIndentBody, bodyWidth, pageWidth, delta As Single
        Dim leftIndentHeader As Single
        '
        myDoc = tbl.Range.Document
        splitPara = Nothing
        para = Nothing
        '
        numHeaderRows = 0
        leftIndent = 0
        leftIndentHeader = 0
        cellPadding = 0
        leftIndentBody = 0
        bodyWidth = 0
        '
        pageWidth = 0
        delta = 0
        '
        tblTop = Nothing
        dr = Nothing
        drCell = Nothing
        '
        strCaption = ""

        'Get Table properties..If we are successful then we can operate on it
        If objTablesMgr.tbl_get_tableProperties(tbl, numHeaderRows, leftIndent, cellPadding, leftIndentBody, bodyWidth) Then
            Try
                pageWidth = objGlobals.glb_get_widthBetweenMargins(tbl.Range.Sections.Item(1))
                delta = pageWidth - bodyWidth
                '
                'All WCAG tables must be inline, then split the table to egt the header row(s) in tblTop
                objTablesMgr.tbl_convert_toInLine(tbl)
                '
                dr = tbl.Rows.Item(numHeaderRows + 1)
                objTablesMgr.tbl_split_Table(dr, tbl, splitPara, tblTop)
                '
                If bodyWidth <= pageWidth Then leftIndentBody = 0.0
                leftIndentHeader = leftIndentBody + leftIndent
                objTablesMgr.tbl_convert_oneRowTableToWCAG(cellPadding, leftIndentHeader, tblTop, "aac Table (tblHeader)")
                '
                splitPara.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                splitPara.Format.SpaceBefore = 0.0
                splitPara.Format.SpaceAfter = 0.0
                splitPara.Range.Font.Size = 1.0
                splitPara.KeepWithNext = True
                '
                tbl.Style = myDoc.Styles.Item("aac Table (with lines)")
                '
                If bodyWidth <= pageWidth Then
                    For Each dr In tbl.Rows
                        dr.LeftIndent = 0.0
                    Next
                Else
                    For Each dr In tbl.Rows
                        dr.LeftIndent = delta
                        'dr.LeftIndent = leftIndentBody + leftIndent
                    Next
                End If
                '
                'Get the Source Rows and do a pre-emptive merge to ensure that we
                'get consistent behaviour
                '
                If objTablesMgr.tbl_find_tableBodyBottom(dr, tbl) Then
                    objTablesMgr.tbl_split_Table(dr, tbl, splitPara, tblTop)
                    '
                    For Each dr In tbl.Rows
                        dr.Cells.Merge()
                    Next
                    '
                    rng = tbl.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
                    '
                    rng.ParagraphFormat.LeftIndent = leftIndentBody
                    splitPara.Range.Delete()
                    '
                End If
                GoTo finis

            Catch ex As Exception

            End Try
            Try
                pageWidth = objGlobals.glb_get_widthBetweenMargins(tbl.Range.Sections.Item(1))
                delta = pageWidth - bodyWidth
                '
                'All WCAG tables must be inline
                objTablesMgr.tbl_convert_toInLine(tbl)
                If numHeaderRows >= 1 Then
                    '
                    'Remove padding and indents... If we apply the basic style at this point with
                    'the indents in place, then entire table is widened to match the first row
                    '
                    For i = 1 To numHeaderRows
                        dr = tbl.Rows.Item(i)
                        drCell = dr.Cells.Item(1)
                        drCell.LeftPadding = 0.0
                        drCell.Width = drCell.Width + leftIndent        'remember leftIndent is negative
                    Next
                    '
                    If bodyWidth <= pageWidth Then
                        For Each dr In tbl.Rows
                            dr.LeftIndent = 0.0
                        Next
                    Else
                        For Each dr In tbl.Rows
                            dr.LeftIndent = delta
                        Next
                    End If
                    '
                    'This forces the table back to a regular structure as defined by the Style
                    tbl.Style = myDoc.Styles.Item("aac Table (Basic)")
                    tbl.ApplyStyleHeadingRows = True
                    '
                    'Re-Colour the Header Rows
                    For i = 1 To numHeaderRows
                        dr = tbl.Rows.Item(i)
                        dr.Shading.ForegroundPatternColor = Me._wcag_color_backcolour_old
                        dr.Shading.BackgroundPatternColor = Me._wcag_color_backcolour_old
                        '
                    Next
                    '
                    'Gte the caption paragraph which is assumed to be directly above the table
                    strCaption = objTablesMgr.tbl_get_tblCaption(tbl, para)
                    '
                    'Now set the indents and adjust the indent of the Caption
                    If bodyWidth <= pageWidth Then
                        For Each dr In tbl.Rows
                            dr.LeftIndent = 0.0
                            objParaMgr.paras_set_HangingIndent(0.0, para)
                        Next
                    Else
                        For Each dr In tbl.Rows
                            dr.LeftIndent = delta
                            objParaMgr.paras_set_HangingIndent(delta, para)
                        Next
                    End If
                    '
                    'Get the Source Rows and do a pre-emptive merge to ensure that we
                    'get consistent behaviour
                    '
                    If objTablesMgr.tbl_find_tableBodyBottom(dr, tbl) Then
                        objTablesMgr.tbl_split_Table(dr, tbl, splitPara, tblTop)
                        For Each dr In tbl.Rows
                            dr.Cells.Merge()
                        Next
                        rng = tbl.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
                        rng.ParagraphFormat.LeftIndent = leftIndentBody
                        splitPara.Range.Delete()
                        '
                    End If

                End If


            Catch ex As Exception

            End Try
        End If
        '
finis:
    End Sub
    '
    '
    ''' <summary>
    ''' This method will convert and adjust the selected Table Based banner Header
    ''' to meet WCGA Requirements.. The Headig Level 1 is placed just below the banner
    ''' whilst the Banner Style is chnaged
    ''' paragraph version
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub wcag_convertbanner_toTableVersion(ByRef tbl As Word.Table, strBannerTag As String)
        Dim myDoc As Word.Document
        Dim sect As Word.Section
        Dim drCell As Word.Cell
        Dim objTblsMgr As cTablesMgr
        Dim tblTopPart, tblGlossary As Word.Table
        Dim rng As Word.Range
        Dim dr As Word.Row
        Dim strChptTitle As String
        Dim strStyleName As String
        Dim para As Word.Paragraph
        Dim styl, stylGloss As Word.Style
        Dim hasPageBreakBefore As Boolean
        '
        myDoc = tbl.Range.Document
        sect = tbl.Range.Sections.Item(1)
        styl = myDoc.Styles.Item(strBannerTag)
        objTblsMgr = New cTablesMgr(myDoc)
        '
        'Get the pageBreakBefore setting for this situation, so that we can re-establish
        'it after we have finished
        hasPageBreakBefore = styl.ParagraphFormat.PageBreakBefore
        '
        strStyleName = "Heading 1"
        '
        'Check if the banner has already been processed (e.g. someone is re running the conversion
        'to wcag.. If so we don't want to reprocess an already processed banner
        drCell = tbl.Range.Cells.Item(3)
        rng = drCell.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        styl = rng.Style
        If styl.NameLocal = "Heading 1 (wcag)" Then Exit Sub

        '
        Select Case strBannerTag
            Case "tag_execBanner"
                strStyleName = "Heading 1 (ES)"
            Case "tag_chapterBanner"
                strStyleName = "Heading 1"
            Case "tag_appendixChapter"
                strStyleName = "Heading 1 (AP)"
            Case "tag_glossary_Chpt", "tag_biblio_Chpt", "tag_refs_Chpt"
                strStyleName = "Heading (glossary)"
                '
                'Now let's apply the "aac Table (no lines)" style to any tables used in the
                'glossary
                For Each tblGlossary In tbl.Range.Sections.Item(1).Range.Tables
                    rng = tblGlossary.Range.Cells.Item(1).Range
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    stylGloss = rng.Style
                    If stylGloss.NameLocal = "Body Text" Or stylGloss.NameLocal = "Normal" Or stylGloss.NameLocal = "Normal - no space" Then
                        tblGlossary.Style = myDoc.Styles.Item("aac Table (no lines)")
                        tblGlossary.ApplyStyleHeadingRows = True
                    End If
                Next
            Case "tag_partBanner", "tag_appendixPart"

        End Select
        '
        '
        'Remove the colour graphic
        drCell = tbl.Range.Cells.Item(1)
        rng = drCell.Range
        rng.Delete()
        '
        rng.ParagraphFormat.PageBreakBefore = False
        '
        tbl.Style = myDoc.Styles.Item("aac Chapter Banner")
        'tbl.Style = myDoc.Styles.Item("aac Table (no lines)")

        tbl.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
        tbl.Range.Font.Color = RGB(255, 254, 255)
        tbl.Shading.ForegroundPatternColor = RGB(54, 31, 76)
        '
        tbl.Rows.Alignment = WdRowAlignment.wdAlignRowLeft
        'tbl.Range.ParagraphFormat.Alignment = WdRowAlignment.wdAlignRowLeft

        '
        drCell = tbl.Range.Cells.Item(3)
        rng = drCell.Range
        strChptTitle = rng.Text
        '
        'Remove paragraph mark at the end
        strChptTitle = Left(strChptTitle, Len(strChptTitle) - 2)
        '
        'Now add the Heading at the bottom of the Table
        'rng = tbl.Range
        'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        'rng.Text = strChptTitle
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'para = rng.Paragraphs.First
        'para.Style = myDoc.Styles.Item(strStyleName)
        'para.Range.Font.Size = 1
        'para.Range.Font.Color = RGB(0, 0, 0)
        '
        '
        'Now add the Heading at the top of the Table, then split the table with
        'a function that return the results of the split. That is, the top table part
        'the bottom table part and the paragraph between the tables
        '
        tbl.Rows.Add(tbl.Rows.First)
        sect = tbl.Range.Sections.Item(1)
        para = Nothing
        tblTopPart = Nothing
        '
        If objTblsMgr.tbl_split_Table(2, tbl, para, tblTopPart) Then
            tblTopPart.Delete()
            '
            rng = para.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Text = strChptTitle
            '
            para = rng.Paragraphs.First
            '
            styl = myDoc.Styles.Item(strStyleName)
            styl.ParagraphFormat.PageBreakBefore = hasPageBreakBefore
            styl.ParagraphFormat.SpaceBefore = 0
            styl.ParagraphFormat.SpaceAfter = 0
            '
            para.Style = styl
            para.Range.Font.Size = 1
            para.Range.Font.Color = RGB(0, 0, 0)
            '
            drCell = tbl.Range.Cells.Item(3)
            rng = drCell.Range
            rng.Delete()
            rng.Style = myDoc.Styles.Item("Heading 1 (wcag)")
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Text = strChptTitle
            '
            Try
                dr = tbl.Rows.Last
                dr.Delete()
            Catch ex As Exception

            End Try
        End If

    End Sub
    '
    ''' <summary>
    ''' This method will convert the selected banner Header to an numLines line back coloured
    ''' paragraph version
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub wcag_convertbanner_toParaVersion(ByRef tbl As Word.Table, numLines As Integer)
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        Dim brdr As Word.Border
        Dim i As Integer
        '
        drCell = tbl.Range.Cells.Item(1)
        rng = drCell.Range
        rng.Delete()
        '
        tbl.Rows.Item(3).Delete()
        tbl.Rows.Item(1).Delete()

        rng = tbl.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
        rng.Shading.BackgroundPatternColor = RGB(54, 31, 76)
        rng.Paragraphs.Last.Range.Delete()
        brdr = rng.ParagraphFormat.Borders.Item(WdBorderType.wdBorderBottom)
        brdr.Color = RGB(54, 31, 76)
        brdr.LineStyle = WdLineStyle.wdLineStyleSingle
        brdr.LineWidth = WdLineWidth.wdLineWidth025pt
        rng.ParagraphFormat.Borders.DistanceFromBottom = 10
        '
        'Now add Soft Returns
        For i = 1 To numLines
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Text = vbVerticalTab
        Next i


    End Sub
    '
    '
    ''' <summary>
    ''' This method will convert all Boxes, KeyFindings and Recommendations in the document myDoc
    ''' to Accessibility compliant forms
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub wcag_convert_aacBoxesToWCAG(ByRef myDoc As Word.Document)
        Dim tbl As Word.Table
        '
        'Add the Box Table Style if it doesn't already exist
        'Me.wcag_stylesTable_addAACBox(myDoc)
        '
        For Each tbl In myDoc.Tables
            Me.wcag_convert_aacBoxToWCAG(tbl)
        Next
        '
    End Sub
    ''' <summary>
    ''' This method expects as input a Box, KeyFinding or Recommendation table. it will convert
    ''' all of these to Accessibility compliant constructs.
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub wcag_convert_aacBoxToWCAG(ByRef tbl As Word.Table)
        Dim myDoc As Word.Document
        Dim objTablesMgr As cTablesMgr
        Dim tblTop As Word.Table
        Dim splitPara As Word.Paragraph
        Dim rng As Word.Range
        '
        myDoc = tbl.Range.Document
        '
        objTablesMgr = New cTablesMgr(myDoc)
        splitPara = Nothing
        tblTop = Nothing
        '
        Try
            If objTablesMgr.tbl_is_AACBox(tbl) Then
                If objTablesMgr.tbl_split_Table(2, tbl, splitPara, tblTop) Then
                    rng = tblTop.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
                    splitPara.Range.Delete()
                    '
                    If objTablesMgr.tbl_split_Table(2, tbl, splitPara, tblTop) Then
                        rng = tbl.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
                        splitPara.Range.Delete()
                        '
                        tblTop.ApplyStyleHeadingRows = True
                        tblTop.Style = myDoc.Styles.Item("aac Table (Box)")
                    End If
                    '
                    '
                End If

            End If
        Catch ex As Exception

        End Try

    End Sub
    '
    '
    ''' <summary>
    ''' This method will copy the caption text into all of the inLine Shapes contained
    ''' in the range rng
    ''' </summary>
    ''' <param name="captionParagraph"></param>
    ''' <param name="rng"></param>
    Public Sub wcag_alttext_write(ByRef captionParagraph As Word.Paragraph, rng As Word.Range)
        Dim strAltText As String
        Dim strTokens As String()
        '
        Try
            strAltText = captionParagraph.Range.Text
            strTokens = strAltText.Split(vbTab)
            strAltText = strTokens(1).Trim()
        Catch ex2 As Exception
            strAltText = ""
        End Try
        '
        '
        If rng.InlineShapes.Count > 0 Then
            For Each iShp In rng.InlineShapes
                If iShp.AlternativeText Like "*Description automatically generated" Or iShp.AlternativeText = "" Then
                    iShp.AlternativeText = strAltText
                End If
            Next
        End If

    End Sub
    '
    '
    ''' <summary>
    ''' This method will write the specified text (strAlttext) onto iShp
    ''' </summary>
    ''' <param name="strAltText"></param>
    ''' <param name="iShp"></param>
    Public Sub wcag_alttext_write(strAltText As String, ByRef iShp As InlineShape)
        '
        Try
            iShp.AlternativeText = strAltText
        Catch ex2 As Exception

        End Try
        '

    End Sub

    '
    '
End Class
