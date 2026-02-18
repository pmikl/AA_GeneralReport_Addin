Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports System.Drawing
Public Class cChptBanner
    Inherits cGlobals
    Public tag_letter, tag_memo, tag_brief, tag_coverPage, tag_cont_Front, tag_toc, tag_es, tag_chpt_body As String
    Public tag_div, tag_glos, tag_glos_bib, tag_glos_refsCited, tag_glos_wrks As String
    Public tag_divAP, tag_chpt_AP, tag_cont_Back As String
    '
    Public strBannerType As String              'Specifies whetehr the banner is an image stored in Resources, or an rgb filled shape, or a shape filled with the stored banner image
    Public rgbFill As Long                      'Fill colour to be used if the banner is a Shape
    '
    Public Banner_Std_Image As Image
    Public Banner_Std_Shape As Word.Shape
    '
    Public objWCAG As cWCAGMgr

    ''' <summary>
    ''' Allowed values for strBannerType are 'image', 'rectangle', 'roundedRectangle'. The default fill for
    ''' Shapes is the standard background purple
    ''' </summary>
    ''' <param name="strBannerType"></param>
    Public Sub New(Optional strBannerType As String = "image")
        MyBase.New()
        '
        Dim objFileMgr As New cFileHandler()
        Dim objRsrcsMgr As New cResourcesMgr()

        Me.tag_letter = "tag_letter"
        Me.tag_memo = "tag_memo"
        Me.tag_brief = "tag_aaBrief"
        Me.tag_coverPage = "cp"
        Me.tag_cont_Front = "contacts_Front"
        Me.tag_toc = "toc"
        Me.tag_es = "ES"
        Me.tag_chpt_body = "body"
        Me.tag_div = "div"
        Me.tag_glos = "glos"
        Me.tag_glos_bib = "bib"
        Me.tag_glos_refsCited = "refs"
        Me.tag_glos_wrks = "wrks"
        Me.tag_divAP = "divAP"
        Me.tag_chpt_AP = "AP"
        Me.tag_cont_Back = "contacts_Back"
        '
        '*** Insert from Resources
        Me.Banner_Std_Image = objRsrcsMgr.rsrcs_get_bannerImage()
        'Me.Banner_Std_Image = My.Resources.banner_KI_Sunset_03
        'Me.Banner_Std_Image = My.Resources.banner_Nasa_bg_image

        Me.Banner_Std_Shape = Nothing
        '
        'Me.strBannerType = strBannerType
        'Me.strBannerType = "rectangle"
        'Me.strBannerType = "roundedRectangle"
        'Me.strBannerType = "flowChartAltProcess"
        'Me.strBannerType = "rectangleSnipRound"
        '
        'Me.strBannerType = "rectangleImageFilled"
        'Me.strBannerType = "roundedRectangleImageFilled"
        'Me.strBannerType = "flowChartAltProcessImageFilled"
        'Me.strBannerType = "rectangleSnipRoundImageFilled"
        Me.strBannerType = "rectangle"
        '
        Me.rgbFill = _glb_colour_purple_Dark
        '
        Me.objWCAG = New cWCAGMgr()

    End Sub
    '
    Public Function bnr_get_tagStyles(strBannerType As String) As String
        Dim strTagStyle As String
        '
        strTagStyle = "tag_chapter_Banner"
        '
        Select Case strBannerType
            Case Me.tag_letter
                strTagStyle = "tag_aa_stn_letter"
            Case Me.tag_memo
                strTagStyle = "tag_aa_stn_memo"
            Case Me.tag_brief
                strTagStyle = "tag_aaBrief"
            Case Me.tag_coverPage
                strTagStyle = "tag_coverPage"
            Case Me.tag_cont_Front
                strTagStyle = "tag_contactsPage-Front"
            Case Me.tag_toc
                'strTagStyle = "TOC Heading"
                strTagStyle = "tag_toc"
            Case Me.tag_es
                strTagStyle = "tag_execBanner"
            Case Me.tag_chpt_body
                strTagStyle = "tag_chapterBanner"
            Case Me.tag_div
                strTagStyle = "tag_partBanner"
            Case Me.tag_glos
                strTagStyle = "tag_glossary_Chpt"
            Case Me.tag_glos_bib
                strTagStyle = "tag_biblio_Chpt"
            Case Me.tag_glos_refsCited
                strTagStyle = "tag_refs_Chpt"
            Case Me.tag_glos_wrks
                strTagStyle = "tag_worksCited_Chpt"
            Case Me.tag_divAP
                strTagStyle = "tag_appendixPart"
            Case Me.tag_chpt_AP
                strTagStyle = "tag_appendixChapter"
            Case Me.tag_cont_Back
                strTagStyle = "tag_contactsPage-Back"
                '
        End Select
        '
        Return strTagStyle
        '
    End Function
    '
    ''' <summary>
    ''' This method will accept a table (tbl) and if that table has a tag style of
    ''' 'tag_ExecBanner', 'tag_chapterBanner', 'tag_appendixChapter' it will convert
    ''' that banner to either a standard banner (strBannerType = "toLarge") or to the smaller
    ''' short report banner (strBannerType = "toSmall")
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="strBannerType"></param>
    ''' <returns></returns>
    Public Function bnr_resize_BannerBase(ByRef tbl As Word.Table, strBannerType As String) As Word.Table
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        Dim myStyle As Word.Style
        '
        drCell = tbl.Range.Cells.Item(1)
        rng = drCell.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        myStyle = rng.Style
        '
        If myStyle.NameLocal = "tag_execBanner" Or myStyle.NameLocal = "tag_chapterBanner" Or myStyle.NameLocal = "tag_appendixChapter" Then
            Select Case strBannerType
                Case "toSmall"
                    tbl.Rows.Item(1).Height = 5.0
                    tbl.Rows.Item(2).Height = 92.0
                    rng.Paragraphs.Item(1).PageBreakBefore = False
                    '
                    Me.bnr_resize_image(tbl)
                Case "toLarge"
                    tbl.Rows.Item(1).Height = 19.1
                    tbl.Rows.Item(2).Height = 138.0
                    rng.Paragraphs.Item(1).PageBreakBefore = True
                    '
                    Me.bnr_resize_image(tbl)
            End Select
        End If
        '
        Return tbl

    End Function
    '
    Public Function bnr_is_Chapter_Bdy_or_ES_or_AP(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        Dim myStyle As Word.Style
        '
        rslt = False
        drCell = tbl.Range.Cells.Item(1)
        rng = drCell.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        myStyle = rng.Style
        '
        If myStyle.NameLocal = "tag_execBanner" Or myStyle.NameLocal = "tag_chapterBanner" Or myStyle.NameLocal = "tag_appendixChapter" Then
            rslt = True
        End If
        '
        Return rslt
    End Function
    '
    '
    Public Function bnr_is_Chapter(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        Dim myStyle As Word.Style
        '
        rslt = False
        drCell = tbl.Range.Cells.Item(1)
        rng = drCell.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        myStyle = rng.Style
        '
        If myStyle.NameLocal = "tag_chapterBanner" Then
            rslt = True
        End If
        '
        Return rslt
    End Function
    '

    '
    Public Function bnr_is_ChapterES(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        Dim myStyle As Word.Style
        '
        rslt = False
        drCell = tbl.Range.Cells.Item(1)
        rng = drCell.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        myStyle = rng.Style
        '
        If myStyle.NameLocal = "tag_execBanner" Then
            rslt = True
        End If
        '
        Return rslt
    End Function
    '
    Public Function bnr_is_ChapterAP(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        Dim myStyle As Word.Style
        '
        rslt = False
        drCell = tbl.Range.Cells.Item(1)
        rng = drCell.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        myStyle = rng.Style
        '
        If myStyle.NameLocal = "tag_appendixChapter" Then
            rslt = True
        End If
        '
        Return rslt
    End Function


    '
    ''' <summary>
    ''' This method will insert at the specified range
    ''' </summary>
    ''' <param name="doBannerImage"></param>
    ''' <returns></returns>
    Public Function bnr_insert_BannerBase(ByRef rng As Word.Range, doBannerImage As Boolean, Optional strRptMode As String = "", Optional lstOfBannerSettings As Collection = Nothing) As Word.Table
        Dim tbl As Word.Table
        Dim sect As Word.Section
        Dim rngLocal As Word.Range
        Dim bannerWidth As Single
        Dim objBBMgr As New cBBlocksHandler()
        Dim objChptBnr As New cChptBanner()
        Dim objRptMgr As New cReport()
        Dim oldPictWrapType As WdWrapTypeMerged
        'Dim imgBanner As Image
        Dim objFileMgr As New cFileHandler()
        Dim doBanners As Boolean
        Dim strBannerType As String
        '
        If IsNothing(lstOfBannerSettings) Then lstOfBannerSettings = Me.bnr_get_BannerSettings(Me.tag_chpt_body)
        If strRptMode = "" Then strRptMode = objRptMgr.rpt_isPrt
        '
        '
        'strTagStyle = CStr(lstOfBannerSettings("strTagStyle"))
        'strHeadingStyle = CStr(lstOfBannerSettings("strHeadingStyle"))
        'strChptNumberStyle = CStr(lstOfBannerSettings("strChptNumberStyle"))
        'strHeadingText = CStr(lstOfBannerSettings("strHeadingText"))
        'strSequenceId = CStr(lstOfBannerSettings("strSequenceId"))
        '
        '
        '*** The collection generated by Me.bnr_get_BannerSettings has amongst its parameters
        '*** a value for strBannerType which can be 'Me.sectType_cp, Me.sectType_cont_Front, Me.sectType_toc,
        '*** Me.sectType_ES, Me.sectType_body, Me.sectType_div, Me.sectType_glos, Me.sectType_bib, Me.sectType_refs, Me.sectType_wrks,
        '*** Me.sectType_divAP, Me.sectType_AP, Me.sectType_cont_Back
        '
        doBanners = Me._glb_doBanners_is_On
        strBannerType = lstOfBannerSettings("strBannerType")
        '
        tbl = Nothing
        sect = rng.Sections.Item(1)
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        '
        doBanners = False
        Select Case strBannerType
            Case Me.tag_div, Me.tag_divAP
                doBanners = False
        End Select
        '
        If doBanners Then
            bannerWidth = glb_get_widthBetweenMargins(sect)
            tbl = rng.Tables.Add(rng, 3, 2)
            '
            Try
                'tbl.Style = rng.Document.Styles("Table Grid")
            Catch ex As Exception

            End Try
            '
            tbl.Borders.InsideLineStyle = WdLineStyle.wdLineStyleNone
            tbl.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone
            '
            'Set the bottom spacer row
            tbl.Rows.Item(3).HeightRule = WdRowHeightRule.wdRowHeightExactly
            tbl.Rows.Item(3).Height = 11.4
            tbl.Rows.Item(3).Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("spacer")
            '
            tbl.Rows.Item(1).Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("spacer")
            tbl.TopPadding = 0.0
            tbl.BottomPadding = 0.0
            tbl.LeftPadding = 0.0
            tbl.RightPadding = 0.0
            '
            tbl.Range.Cells.Item(4).RightPadding = 11.2
            '
            'Fix the column 2 width (i.e. the Chapter Number Column) and then adjust the 
            'width of column 1
            tbl.Columns.Item(2).Width = 118.3
            tbl.Columns.Item(1).Width = bannerWidth - tbl.Columns.Item(2).Width
            '
            tbl.Rows.Item(1).HeightRule = WdRowHeightRule.wdRowHeightExactly
            tbl.Rows.Item(2).HeightRule = WdRowHeightRule.wdRowHeightExactly
            '
            tbl.Rows.Item(1).Height = 19.1
            tbl.Rows.Item(2).Height = 138.0
            '
            'Select Case strRptMode
            'Case objRptMgr.modeLong
            'tbl.Rows.Item(1).Height = 19.1
            'tbl.Rows.Item(2).Height = 138.0
            'Case objRptMgr.modeLongLandscape
            'tbl.Rows.Item(1).Height = 19.1
            'tbl.Rows.Item(2).Height = 138.0
            'tbl.Rows.LeftIndent = -12.4
            'tbl.Columns.Item(1).Width = tbl.Columns.Item(1).Width + 12.4
            'Case objRptMgr.modeShort
            'tbl.Rows.Item(1).Height = 5.0
            'tbl.Rows.Item(2).Height = 92.0
            'End Select
            '
            tbl.Range.Cells.Item(3).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
            tbl.Range.Cells.Item(4).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
            '
            'Get the Heading style and adjust the width of the column to allow for Heading level
            'indents and then adjust the Table indent
            'tbl.Range.Cells.Item(3).Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item(strHeadingStyle)
            '
            If doBannerImage Then
                'In some instances we may not want the image... If we do, just make sure that the default
                'will insert a Floating picture behind the text... This doesn't matter if the stored
                'Building Block is stored as a "Floating picture behind the text", but if we ever replace this 
                'with a software construct it might matter
                '
                oldPictWrapType = objRptMgr.objGlobals.glb_get_wrdApp.Options.PictureWrapType
                objRptMgr.objGlobals.glb_get_wrdApp.Options.PictureWrapType = WdWrapTypeMerged.wdWrapMergeBehind
                '
                rngLocal = tbl.Range.Cells.Item(1).Range
                rngLocal.Collapse(WdCollapseDirection.wdCollapseStart)
                '
                '*** INsert from Building Blocks or from Resources.. Eventually we must go to resources to
                '*** decouple the software from the Template. Me.Banner_Std_Image is the standard banner, which
                '*** is initialised in Public Sub New()... Note that rngLocal in the Resources option is passed
                '*** by reference and modified in the routine
                '
                'rngLocal = objBBMgr.insertBuildingBlockFromDefaultLibToRange("aac_Img_chptBanner_std", "Images", rngLocal)
                '
                '40 is just a dummy, the resize happens elsewhere
                'Me.Banner_Std_Shape = objFileMgr.file_insert_imageFromResources2(rngLocal, 40.0, Me.Banner_Std_Image)
                '
                'Me.Banner_Std_Shape = Me.bnr_insert_BannerImage(tbl, Me.Banner_Std_Image)
                'Me.bnr_insert_BannerBackground.
                Me.rgbFill = RGB(255, 255, 255)
                Me.Banner_Std_Shape = Me.bnr_insert_BannerBackground(tbl, Me.strBannerType)
                '
                '*** Set the banner to white
                '
                'Me.Banner_Std_Shape.Fill.Solid()
                'Me.Banner_Std_Shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
                'Me.Banner_Std_Shape.Line.Visible = False
                'Me.Banner_Std_Shape = Me.bnr_insert_BannerShape(tbl)


                '*** To speed this up, once we have one banner, then don't have to retrieve it again
                '
                '*** Code to insert Shape instead of image Building Block to go here
                'Dim objTestMgr As New cTestingMgr()
                'objTestMgr.insert_ChapterBanner(rngLocal)
                '
                '
                'tbl = objChptBnr.bnr_resize_image(rngLocal)
                'tbl = Me.bnr_resize_image(Me.Banner_Std_Shape)
                '
                objRptMgr.objGlobals.glb_get_wrdApp.Options.PictureWrapType = oldPictWrapType
                '
            End If
            '
            Me.bnr_format_Banner(tbl, lstOfBannerSettings)
            '
        Else
            Me.bnr_format_Banner(rng, lstOfBannerSettings)
        End If
        '

finis:
        Return tbl
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method is meant to be an alternative to file_insert_imageFromFile. In this case we insert any image (img)
    ''' supplied and return it as an inLine SHape. Typically img is obtained form My.Resources. So we have no reliance
    ''' on external file structures
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="bannerImg"></param>
    ''' <returns></returns>
    Public Function bnr_insert_BannerImage(ByRef rng As Word.Range, bannerImg As Image) As Word.Shape
        Dim objImgMgr As New cImageMgr()
        Dim shp As Word.Shape
        Dim j As Integer
        shp = Nothing
        '
        j = 0
        System.Windows.Forms.Clipboard.SetImage(bannerImg)
        '
        Try
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            '*** The inLine option doesn't seem to have any affect 20231124. SO we have to convert to an inline shape ourselves
            rng.PasteSpecial(DataType:=WdPasteDataType.wdPasteBitmap, Placement:=WdOLEPlacement.wdFloatOverText)
            rng.MoveEnd(WdUnits.wdParagraph, 1)
            '
            shp = objImgMgr.img_get_ImageAsShape(rng)
            Me.bnr_resize_image(shp)
            'shp.LockAspectRatio = True
            'shp.Width = width
            '
            'rng.MoveEnd(WdUnits.wdParagraph, -1)
            '
        Catch ex As Exception
            shp = Nothing
        End Try
        '
        Return shp


    End Function
    '
    ''' <summary>
    ''' This method will insert an image (from Me.Banner_Std_Image) or a shape as the banner background. Values for
    ''' strBackGroundType are; 'image', 'rectangle', 'roundedRectangle'
    ''' </summary>
    ''' <param name="BannerTbl"></param>
    ''' <param name="strBackGroundType"></param>
    ''' <returns></returns>
    Public Function bnr_insert_BannerBackground(ByRef BannerTbl As Word.Table, Optional strBackGroundType As String = "rectangle") As Word.Shape
        Dim objImgMgr As New cImageMgr()
        Dim img As Image
        Dim drCell As Word.Cell
        Dim myDoc As Word.Document
        Dim rng As Word.Range
        Dim shp As Word.Shape
        shp = Nothing
        '
        myDoc = Me.glb_get_wrdActiveDoc
        '
        Select Case strBackGroundType
            Case "image"
                System.Windows.Forms.Clipboard.SetImage(Me.Banner_Std_Image)
                drCell = BannerTbl.Range.Cells.Item(1)
                rng = drCell.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.PasteSpecial(DataType:=WdPasteDataType.wdPasteBitmap, Placement:=WdOLEPlacement.wdFloatOverText)
                rng.MoveEnd(WdUnits.wdParagraph, 1)
                '
                'shp = objImgMgr.img_get_ImageAsShape(rng)
                shp = objImgMgr.img_get_ImageAsShape(drCell)
                '
                Me.bnr_resize_image(shp)
                '
                System.Windows.Forms.Clipboard.Clear()
            '
            Case "rectangle", "roundedRectangle", "flowChartAltProcess", "rectangleSnipRound"
                drCell = BannerTbl.Range.Cells.Item(1)
                rng = drCell.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                '
                If strBackGroundType = "rectangle" Then shp = myDoc.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 120, 40, rng)
                If strBackGroundType = "roundedRectangle" Then shp = myDoc.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 0, 0, 120, 40, rng)
                If strBackGroundType = "flowChartAltProcess" Then shp = myDoc.Shapes.AddShape(MsoAutoShapeType.msoShapeFlowchartAlternateProcess, 0, 0, 120, 40, rng)
                If strBackGroundType = "rectangleSnipRound" Then shp = myDoc.Shapes.AddShape(MsoAutoShapeType.msoShapeSnipRoundRectangle, 0, 0, 120, 40, rng)
                '
                Me.bnr_resize_image(shp)
                '
                shp.Fill.Solid()
                shp.Fill.BackColor.RGB = Me.rgbFill
                shp.Fill.ForeColor.RGB = Me.rgbFill
                shp.Line.Visible = False
                '
                '
            Case "rectangleImageFilled", "rectangleSnipRoundImageFilled"
                drCell = BannerTbl.Range.Cells.Item(1)
                rng = drCell.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                '
                If strBackGroundType = "rectangleImageFilled" Then shp = myDoc.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 120, 40, rng)
                If strBackGroundType = "roundedRectangleImageFilled" Then shp = myDoc.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 0, 0, 120, 40, rng)
                If strBackGroundType = "flowChartAltProcessImageFilled" Then shp = myDoc.Shapes.AddShape(MsoAutoShapeType.msoShapeFlowchartAlternateProcess, 0, 0, 120, 40, rng)
                If strBackGroundType = "rectangleSnipRoundImageFilled" Then shp = myDoc.Shapes.AddShape(MsoAutoShapeType.msoShapeSnipRoundRectangle, 0, 0, 120, 40, rng)
                '
                Me.bnr_resize_image(shp)
                '
                img = Me.Banner_Std_Image
                System.Windows.Forms.Clipboard.SetImage(img)

                Dim strFilePath As String
                strFilePath = My.Computer.FileSystem.SpecialDirectories.MyPictures + "\aac_banner_file.jpg"
                '
                img.Save(strFilePath)
                shp.Fill.UserPicture(strFilePath)
                '
                System.Windows.Forms.Clipboard.Clear()

                'shp.Fill.Solid()
                'shp.Fill.BackColor.RGB = Me.rgbFill
                'shp.Fill.ForeColor.RGB = Me.rgbFill
                '
                '

        End Select
        '
        Try
        Catch ex As Exception
            shp = Nothing
        End Try
        '
        Return shp

    End Function
    '    
    Public Function bnr_insert_BannerShape(ByRef BannerTbl As Word.Table) As Word.Shape
        Dim objImgMgr As New cImageMgr()
        Dim drCell As Word.Cell
        Dim myDoc As Word.Document
        Dim rng As Word.Range
        Dim shp As Word.Shape
        shp = Nothing
        '
        myDoc = Me.glb_get_wrdActiveDoc
        '
        Try
            drCell = BannerTbl.Range.Cells.Item(1)
            rng = drCell.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            'shp = myDoc.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 120, 40, rng)
            shp = myDoc.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 0, 0, 120, 40, rng)
            shp.Fill.Solid()
            shp.Fill.BackColor.RGB = Me._glb_colour_purple_Dark
            shp.Fill.ForeColor.RGB = Me._glb_colour_purple_Dark

            'rng.PasteSpecial(DataType:=WdPasteDataType.wdPasteShape, Placement:=WdOLEPlacement.wdFloatOverText)
            'rng.MoveEnd(WdUnits.wdParagraph, 1)
            '
            'shp = objImgMgr.img_get_ImageAsShape(rng)
            Me.bnr_resize_image(shp)
            '
        Catch ex As Exception
            shp = Nothing
        End Try
        '
        Return shp

    End Function
    '

    '
    '
    '
    Public Function bnr_format_Banner(ByRef tbl As Word.Table, ByRef lstOfBannerSettings As Collection) As Word.Table
        Dim strTagStyle, strHeadingStyle, strChptNumberStyle, strHeadingText, strSequenceId, strDoImage As String
        Dim myDoc As Word.Document
        Dim rngLocal As Word.Range
        Dim fld As Word.Field
        '
        myDoc = tbl.Range.Document
        '
        strTagStyle = CStr(lstOfBannerSettings("strTagStyle"))
        strHeadingStyle = CStr(lstOfBannerSettings("strHeadingStyle"))
        strChptNumberStyle = CStr(lstOfBannerSettings("strChptNumberStyle"))
        strHeadingText = CStr(lstOfBannerSettings("strHeadingText"))
        strSequenceId = CStr(lstOfBannerSettings("strSequenceId"))
        strDoImage = CStr(lstOfBannerSettings("strDoImage"))


        tbl.Range.Cells.Item(1).Range.Style = myDoc.Styles.Item(strTagStyle)
        If strChptNumberStyle <> "" Then
            tbl.Range.Cells.Item(4).Range.Style = myDoc.Styles.Item(strChptNumberStyle)
            rngLocal = tbl.Range.Cells.Item(4).Range
            rngLocal.Collapse(WdCollapseDirection.wdCollapseStart)
            fld = rngLocal.Fields.Add(rngLocal, WdFieldType.wdFieldSequence, strSequenceId, False)
            '
        Else
            tbl.Range.Cells.Item(4).Range.Style = myDoc.Styles.Item("spacer")
        End If
        '
        tbl.Range.Cells.Item(3).Range.Style = strHeadingStyle
        tbl.Range.Cells.Item(3).Range.Text = strHeadingText
        '
        If strTagStyle = Me.bnr_get_tagStyles(Me.tag_chpt_AP) Then
            tbl.Range.Cells.Item(2).Range.Style = "Heading 9"
        End If
        '
        Return tbl
        '
    End Function
    '
    Public Function bnr_format_Banner(ByRef rng As Word.Range, ByRef lstOfBannerSettings As Collection) As Word.Range
        Dim strTagStyle, strHeadingStyle, strSubHeadingStyle, strChptNumberStyle, strHeadingText, strSubHeadingText, strSequenceId, strDoImage As String
        Dim myDoc As Word.Document
        Dim para As Word.Paragraph
        '
        myDoc = rng.Document
        '
        strTagStyle = CStr(lstOfBannerSettings("strTagStyle"))
        strHeadingStyle = CStr(lstOfBannerSettings("strHeadingStyle"))
        strSubHeadingStyle = CStr(lstOfBannerSettings("strSubHeadingStyle"))
        strChptNumberStyle = CStr(lstOfBannerSettings("strChptNumberStyle"))
        strHeadingText = CStr(lstOfBannerSettings("strHeadingText"))
        strSubHeadingText = CStr(lstOfBannerSettings("strSubHeadingText"))
        strSequenceId = CStr(lstOfBannerSettings("strSequenceId"))
        strDoImage = CStr(lstOfBannerSettings("strDoImage"))
        '
        Try
            rng.Style = strHeadingStyle
            rng.Text = strHeadingText
            '
            If strSubHeadingText <> "" And strSubHeadingStyle <> "" Then
                Try
                    rng.Style = strHeadingStyle
                    rng.Text = strHeadingText + vbCrLf + strSubHeadingText
                    para = rng.Paragraphs.Item(2)
                    para.Style = glb_get_wrdActiveDoc.Styles.Item(strSubHeadingStyle)
                Catch ex As Exception
                    rng.Style = strHeadingStyle
                    rng.Text = strHeadingText

                End Try
            Else
                rng.Style = strHeadingStyle
                rng.Text = strHeadingText
            End If
            '
            If strTagStyle = Me.bnr_get_tagStyles(Me.tag_chpt_AP) Then
                'tbl.Range.Cells.Item(2).Range.Style = "Heading 9"
            End If

        Catch ex As Exception

        End Try
        '
        Return rng
        '

    End Function
    '
    ''' <summary>
    ''' This method will resize the banner image to fit the banner table (tbl)
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function bnr_resize_image(ByRef tbl As Word.Table) As Word.Shape
        Dim rng As Word.Range
        Dim shp As Word.Shape
        '
        shp = Nothing
        rng = tbl.Range
        Me.bnr_resize_image(rng)
        '
        If tbl.Range.ShapeRange.Count <> 0 Then
            shp = tbl.Range.ShapeRange.Item(1)
        End If
        '
        Return shp
        '
    End Function
    '
    ''' <summary>
    ''' This method will resize the banner image to fit the (first) banner table in the
    ''' rnage rng
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <returns></returns>
    Public Function bnr_resize_image(ByRef rng As Word.Range) As Word.Table
        Dim shp As Word.Shape
        Dim tbl As Word.Table
        '
        shp = Nothing
        tbl = Nothing
        '
        If rng.ShapeRange.Count <> 0 And rng.Tables.Count <> 0 Then
            tbl = rng.Tables.Item(1)
            shp = rng.ShapeRange.Item(1)
            shp.LockAspectRatio = False
            shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
            shp.Top = 0.0
            shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
            shp.Left = 0.0
            shp.ZOrder(MsoZOrderCmd.msoSendBehindText)
            shp.Width = tbl.Columns.Item(1).Width + tbl.Columns.Item(2).Width
            shp.Height = tbl.Rows.Item(1).Height + tbl.Rows.Item(2).Height
            '
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            shp.LockAnchor = True
            '
            objWCAG.wcag_set_decorative(shp, True)
            'shp.LockAspectRatio = True

            '
            'We redo the positioning because the shape can come adrift (almost randomly)
            'so I refix it just to make sure
            'shp.Top = 0.0
            'shp.Left = 0.0
        End If
        '
        Return tbl
        '
    End Function
    '
    Public Function bnr_resize_image(ByRef shp As Word.Shape) As Word.Table
        Dim rng As Word.Range
        Dim tbl As Word.Table
        '
        rng = shp.Anchor
        tbl = rng.Tables.Item(1)
        '
        Try
            shp.LockAspectRatio = False
            shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
            shp.Top = 0.0
            shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
            shp.Left = 0.0
            shp.ZOrder(MsoZOrderCmd.msoSendBehindText)
            shp.Width = tbl.Columns.Item(1).Width + tbl.Columns.Item(2).Width
            shp.Height = tbl.Rows.Item(1).Height + tbl.Rows.Item(2).Height
            '
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            shp.LockAnchor = True
            '
            objWCAG.wcag_set_decorative(shp, True)

        Catch ex As Exception
            tbl = Nothing
        End Try
        '
        Return tbl
    End Function
    ''' <summary>
    ''' This method will get the Banner Settings (styles etc). The return values are dependent on strBannerType
    ''' which can take on the following values; Me.sectType_glos, Me.sectType_es, Me.sectType_body, Me.sectType_AP,
    ''' Me.sectType_bib, Me.sectType_refs, Me.sectType_wrks, Me.sectType_div, Me.sectType_divAP
    ''' 
    ''' </summary>
    ''' <param name="strBannerType"></param>
    ''' <param name="doImage"></param>
    ''' <returns></returns>
    Public Function bnr_get_BannerSettings(strBannerType As String, Optional doImage As Boolean = True) As Collection
        Dim rngTable As Word.Range
        Dim strSequenceId, strHeadingStyle, strSubHeadingStyle, strChptNumberStyle, strHeadingText, strSubHeadingText, strTagStyle, strDoImage As String
        Dim strBannerImage_BBlkName As String
        Dim objBBMgr As New cBBlocksHandler()
        Dim lst As New Collection()
        '
        '
        rngTable = Nothing
        strBannerImage_BBlkName = ""
        '
        strTagStyle = Me.bnr_get_tagStyles(strBannerType)
        '
        strHeadingStyle = ""
        strSubHeadingStyle = ""
        strChptNumberStyle = ""
        strHeadingText = ""
        strSubHeadingText = ""
        strSequenceId = "ChptNum" & " \* ARABIC"

        strDoImage = "true"                                     'Can be used to turn on/off the banner image
        '
        If Not doImage Then
            strDoImage = "false"                                     'Can be used to turn on/off the banner image
        End If
        '
        Try
            Select Case strBannerType
                Case Me.tag_glos
                    strHeadingStyle = "Heading (glossary)"
                    strChptNumberStyle = ""
                    strHeadingText = "Glossary"
                    '
                Case Me.tag_es
                    strHeadingStyle = "Heading 1 (ES)"
                    strChptNumberStyle = ""
                    strHeadingText = "ES Heading 1"
                    '
                Case Me.tag_chpt_body
                    strHeadingStyle = "Heading 1"
                    strChptNumberStyle = "Heading (Chapter)"
                    strHeadingText = "Heading 1"
                    strSequenceId = "ChptNum" & " \* ARABIC"
                    '
                Case Me.tag_chpt_AP
                    strHeadingStyle = "Heading 6"
                    If _glb_doApp_as_HeadingAP Then strHeadingStyle = "Heading 1 (AP)"
                    '
                    strChptNumberStyle = "Heading (Appendix)"
                    strHeadingText = "Appendix Heading 1"
                    strSequenceId = "AppNum" & " \* ALPHABETIC"
                    '
                Case Me.tag_glos_bib
                    strHeadingStyle = "Heading (glossary)"
                    strChptNumberStyle = ""
                    strHeadingText = "Bibliography"
                    '
                Case Me.tag_glos_refsCited
                    strHeadingStyle = "Heading (glossary)"
                    strChptNumberStyle = ""
                    strHeadingText = "References"
                    '
                Case Me.tag_glos_wrks
                    strHeadingStyle = "Heading (glossary)"
                    strChptNumberStyle = ""
                    strHeadingText = "Works cited"
                    '
                Case Me.tag_div
                    strHeadingStyle = "Part - Heading (Banner)"
                    strChptNumberStyle = "Part - Number"
                    'strHeadingText = "Part Number"
                    strHeadingText = "Divider Title"
                    strSequenceId = "NumList" & " \* ROMAN"
                    'strSubHeadingText = "Sub heading"
                    strSubHeadingText = ""
                    strSubHeadingStyle = "Part - SubHead (Banner)"
                    '
                Case Me.tag_divAP
                    strHeadingStyle = "App - Divider (Heading)"
                    strChptNumberStyle = ""
                    strHeadingText = "Appendices"
                    '
            End Select

        Catch ex As Exception

        End Try
        '
        lst.Add(strTagStyle, "strTagStyle")
        lst.Add(strHeadingStyle, "strHeadingStyle")
        lst.Add(strSubHeadingStyle, "strSubHeadingStyle")
        lst.Add(strChptNumberStyle, "strChptNumberStyle")
        lst.Add(strHeadingText, "strHeadingText")
        lst.Add(strSubHeadingText, "strSubHeadingText")
        lst.Add(strSequenceId, "strSequenceId")
        lst.Add(strBannerType, "strBannerType")
        '
        If Not doImage Then strDoImage = "false"
        lst.Add(strDoImage, "strDoImage")
        '
        Return lst
        '
    End Function
    '
    '
    '
    ''' <summary>
    ''' This method will insert at the specified range
    ''' </summary>
    ''' <param name="doBannerImage"></param>
    ''' <returns></returns>
    Public Function xbnr_insert_BannerBase(ByRef rng As Word.Range, doBannerImage As Boolean, strBannerType As String, Optional placeAtSectionStart As Boolean = True) As Word.Table
        Dim tbl As Word.Table
        Dim sect As Word.Section
        Dim rngLocal As Word.Range
        Dim bannerWidth As Single
        Dim objBBMgr As New cBBlocksHandler()
        Dim objRptMgr As New cReport()
        Dim objFldsMgr As New cFieldsMgr()
        Dim lstOfSettings As Collection
        Dim myDoc As Word.Document
        Dim fld As Word.Field

        '
        Dim strTagStyle, strHeadingStyle, strChptNumberStyle, strHeadingText, strSequenceId As String
        '
        lstOfSettings = Me.bnr_get_BannerSettings(strBannerType)
        '
        strTagStyle = CStr(lstOfSettings("strTagStyle"))
        strHeadingStyle = CStr(lstOfSettings("strHeadingStyle"))
        strChptNumberStyle = CStr(lstOfSettings("strChptNumberStyle"))
        strHeadingText = CStr(lstOfSettings("strHeadingText"))
        strSequenceId = CStr(lstOfSettings("strSequenceId"))
        '

        '
        sect = rng.Sections.Item(1)
        myDoc = rng.Document
        '
        If placeAtSectionStart Then
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
        Else

        End If
        '
        '
        bannerWidth = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
        tbl = rng.Tables.Add(rng, 3, 2)
        '
        'Set the bottom spacer row
        tbl.Rows.Item(3).HeightRule = WdRowHeightRule.wdRowHeightExactly
        tbl.Rows.Item(3).Height = 11.4
        tbl.Rows.Item(3).Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("spacer")
        '
        tbl.Rows.Item(1).Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("spacer")
        tbl.TopPadding = 0.0
        tbl.BottomPadding = 0.0
        tbl.LeftPadding = 0.0
        tbl.RightPadding = 0.0
        '
        tbl.Range.Cells.Item(4).RightPadding = 11.2
        '
        'Fix the column 2 width (i.e. the Chapter Number Column) and then adjust the 
        'width of column 1
        tbl.Columns.Item(2).Width = 118.3
        tbl.Columns.Item(1).Width = bannerWidth - tbl.Columns.Item(2).Width
        '
        tbl.Rows.Item(1).HeightRule = WdRowHeightRule.wdRowHeightExactly
        tbl.Rows.Item(2).HeightRule = WdRowHeightRule.wdRowHeightExactly
        '
        tbl.Rows.Item(1).Height = 19.1
        tbl.Rows.Item(2).Height = 138.0

        Select Case strBannerType
            Case MyBase._glb_strBanner_Std
                tbl.Rows.Item(1).Height = 19.1
                tbl.Rows.Item(2).Height = 138.0
            Case MyBase._glb_strBanner_Sht
                tbl.Rows.Item(1).Height = 5.0
                tbl.Rows.Item(2).Height = 92.0
        End Select
        '
        tbl.Range.Cells.Item(3).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
        tbl.Range.Cells.Item(4).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
        '
        'Get the Heading style and adjust the width of the column to allow for Heading level
        'indents and then adjust the Table indent
        'tbl.Range.Cells.Item(3).Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item(strHeadingStyle)
        '
        tbl.Range.Cells.Item(1).Range.Style = myDoc.Styles.Item(strTagStyle)
        'tbl.Range.Cells.Item(4).Range.Style = myDoc.Styles.Item(strChptNumberStyle)
        '
        rngLocal = tbl.Range.Cells.Item(4).Range
        rngLocal.Collapse(WdCollapseDirection.wdCollapseStart)
        fld = rngLocal.Fields.Add(rngLocal, WdFieldType.wdFieldSequence, strSequenceId, False)
        '
        tbl.Range.Cells.Item(3).Range.Style = strHeadingStyle
        tbl.Range.Cells.Item(3).Range.Text = strHeadingText
        '

        '
        If doBannerImage Then
            'In some instances we may not want the image
            rngLocal = tbl.Range.Cells.Item(1).Range
            rngLocal.Collapse(WdCollapseDirection.wdCollapseStart)
            rngLocal = objBBMgr.insertBuildingBlockFromDefaultLibToRange("aac_Img_chptBanner_std", "Images", rngLocal)
            '
            tbl = Me.bnr_resize_image(rngLocal)
            '
        End If
        '

        objFldsMgr.flds_update_SequenceNumbers_Chapters()
finis:
        Return tbl
        '
    End Function
    '
    '


End Class
