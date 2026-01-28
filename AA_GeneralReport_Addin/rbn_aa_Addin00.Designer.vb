Partial Class rbn_aa_Addin00
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(rbn_aa_Addin00))
        Dim RibbonDropDownItemImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl2 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl3 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDialogLauncherImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher = Me.Factory.CreateRibbonDialogLauncher
        Me.tab_aa_Styles = Me.Factory.CreateRibbonTab
        Me.grp_Styles_AAThemes = Me.Factory.CreateRibbonGroup
        Me.mnu_SetTheme = Me.Factory.CreateRibbonMenu
        Me.xbtn__mnuThemes_set_AATheme = Me.Factory.CreateRibbonButton
        Me.xbtn__mnuThemes_set_AAThemeAndStyles = Me.Factory.CreateRibbonButton
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.xbtn_mnuThemes_ActivateTabPGS = Me.Factory.CreateRibbonButton
        Me.xbtn_mnuThemes_PGSToggle = Me.Factory.CreateRibbonButton
        Me.grpStyles_CoverPage = Me.Factory.CreateRibbonGroup
        Me.grpStylesES_StyleSet = Me.Factory.CreateRibbonButton
        Me.grpStylesES_Heading2_ES = Me.Factory.CreateRibbonButton
        Me.grpStylesES_Heading3_ES = Me.Factory.CreateRibbonButton
        Me.grpStylesES_Heading4_ES = Me.Factory.CreateRibbonButton
        Me.grpStylesES_Heading5_ES = Me.Factory.CreateRibbonButton
        Me.grpStyles_Report = Me.Factory.CreateRibbonGroup
        Me.grpStylesRpt_StyleSet = Me.Factory.CreateRibbonButton
        Me.grpStylesRpt_Heading2_Rpt = Me.Factory.CreateRibbonButton
        Me.grpStylesRpt_Heading3_Rpt = Me.Factory.CreateRibbonButton
        Me.grpStylesRpt_Heading4_Rpt = Me.Factory.CreateRibbonButton
        Me.grpStylesRpt_Heading5_Rpt = Me.Factory.CreateRibbonButton
        Me.Separator5 = Me.Factory.CreateRibbonSeparator
        Me.grpStyles_mnu_Heading3Numbering = Me.Factory.CreateRibbonMenu
        Me.grpStyles_mnu_Heading3Numbering_btn_on = Me.Factory.CreateRibbonButton
        Me.grpStyles_mnu_Heading3Numbering_btn_off = Me.Factory.CreateRibbonButton
        Me.grpStyles_NoNum = Me.Factory.CreateRibbonGroup
        Me.grpStylesRpt_HeadingNoNum_StyleSet = Me.Factory.CreateRibbonButton
        Me.grpStylesRpt_Heading2NoNum_Rpt = Me.Factory.CreateRibbonButton
        Me.grpStylesRpt_Heading3NoNum_Rpt = Me.Factory.CreateRibbonButton
        Me.grpStylesRpt_Heading4NoNum_Rpt = Me.Factory.CreateRibbonButton
        Me.grpStylesRpt_Heading5NoNum_Rpt = Me.Factory.CreateRibbonButton
        Me.grpStyles_Appendices = Me.Factory.CreateRibbonGroup
        Me.grpStylesApp_StyleSet = Me.Factory.CreateRibbonButton
        Me.grpStylesApp_Heading1_App = Me.Factory.CreateRibbonButton
        Me.grpStylesApp_Heading2_App = Me.Factory.CreateRibbonButton
        Me.grpStylesApp_Heading3_App = Me.Factory.CreateRibbonButton
        Me.grpStylesApp_Heading4_App = Me.Factory.CreateRibbonButton
        Me.grpStylesApp_Heading5_App = Me.Factory.CreateRibbonButton
        Me.grpStyles_Text = Me.Factory.CreateRibbonGroup
        Me.grpStylesText_BodyText = Me.Factory.CreateRibbonButton
        Me.grpStylesRpt_Intro = Me.Factory.CreateRibbonButton
        Me.grpStylesOther_Quote = Me.Factory.CreateRibbonButton
        Me.grpStylesOther_QuoteBlt = Me.Factory.CreateRibbonButton
        Me.grpStylesOther_QuoteSource = Me.Factory.CreateRibbonButton
        Me.grpStyles_Lists = Me.Factory.CreateRibbonGroup
        Me.grpStylesLists_List1 = Me.Factory.CreateRibbonButton
        Me.grpStylesLists_List2 = Me.Factory.CreateRibbonButton
        Me.grpStylesLists_List3 = Me.Factory.CreateRibbonButton
        Me.Separator55 = Me.Factory.CreateRibbonSeparator
        Me.grpStylesLists_ListNumber1 = Me.Factory.CreateRibbonButton
        Me.grpStylesLists_ListNumber2 = Me.Factory.CreateRibbonButton
        Me.grpStylesLists_ListNumber3 = Me.Factory.CreateRibbonButton
        Me.grpStyles_Emphasis = Me.Factory.CreateRibbonGroup
        Me.tbStyles_mnu_Emphasis = Me.Factory.CreateRibbonMenu
        Me.grpPullouts_emphasisBox_TextStyle_Left_2 = Me.Factory.CreateRibbonButton
        Me.grpPullouts_emphasisBox_TextStyle_Centre_2 = Me.Factory.CreateRibbonButton
        Me.grpPullouts_emphasisBox_TextStyle_Right_2 = Me.Factory.CreateRibbonButton
        Me.grpStyles_resetStyles = Me.Factory.CreateRibbonGroup
        Me.grpStylesTools_to_PrintDefault = Me.Factory.CreateRibbonButton
        Me.grpStylesTools_to_DisplayDefault = Me.Factory.CreateRibbonButton
        Me.tbStyles_grpResetStyles_mnu_ResetStyles = Me.Factory.CreateRibbonMenu
        Me.tabStyles_btn_resetStylesForRptPrt = Me.Factory.CreateRibbonButton
        Me.tabStyles_btn_resetStylesForRptLnd = Me.Factory.CreateRibbonButton
        Me.tabStyles_btn_resetStylesForRptBrf = Me.Factory.CreateRibbonButton
        Me.grpStyles_resetCaptions = Me.Factory.CreateRibbonGroup
        Me.grpStylesTools_resetCaptions = Me.Factory.CreateRibbonButton
        Me.tab_aa_Placeholders = Me.Factory.CreateRibbonTab
        Me.grp_PlaceHolders = Me.Factory.CreateRibbonGroup
        Me.mnuCloseDocuments161 = Me.Factory.CreateRibbonMenu
        Me.grpBoxes_Box = Me.Factory.CreateRibbonButton
        Me.grpBoxes_AppendixBox = Me.Factory.CreateRibbonButton
        Me.grpBoxes_ESBox = Me.Factory.CreateRibbonButton
        Me.grpBoxes_LTBox = Me.Factory.CreateRibbonButton
        Me.Separator31 = Me.Factory.CreateRibbonSeparator
        Me.grpBoxes_CaptionAndHeading = Me.Factory.CreateRibbonButton
        Me.grpBoxes_CaptionAndHeadingES = Me.Factory.CreateRibbonButton
        Me.grpBoxes_CaptionAndHeadingApp = Me.Factory.CreateRibbonButton
        Me.mnuCloseDocuments2233 = Me.Factory.CreateRibbonMenu
        Me.grpBoxes_BoxTextBoldItalic = Me.Factory.CreateRibbonButton
        Me.grpBoxes_BoxText = Me.Factory.CreateRibbonButton
        Me.grpBoxes_SideHeading1 = Me.Factory.CreateRibbonButton
        Me.grpBoxes_SideHeading2 = Me.Factory.CreateRibbonButton
        Me.Separator32 = Me.Factory.CreateRibbonSeparator
        Me.grpBoxes_BoxListBullet = Me.Factory.CreateRibbonButton
        Me.grpBoxes_BoxListBullet2 = Me.Factory.CreateRibbonButton
        Me.grpBoxes_BoxListBullet3 = Me.Factory.CreateRibbonButton
        Me.Separator33 = Me.Factory.CreateRibbonSeparator
        Me.grpBoxes_BoxListNumber = Me.Factory.CreateRibbonButton
        Me.grpBoxes_BoxListNumber2 = Me.Factory.CreateRibbonButton
        Me.grpBoxes_BoxListNumber3 = Me.Factory.CreateRibbonButton
        Me.Separator34 = Me.Factory.CreateRibbonSeparator
        Me.grpBoxes_BoxQuote = Me.Factory.CreateRibbonButton
        Me.grpBoxes_BoxQuoteListBullet = Me.Factory.CreateRibbonButton
        Me.grpBoxes_BoxQuoteSource = Me.Factory.CreateRibbonButton
        Me.Separator35 = Me.Factory.CreateRibbonSeparator
        Me.grpBoxes_boxContent_mnu = Me.Factory.CreateRibbonMenu
        Me.grpBoxes_deleteBoxContent = Me.Factory.CreateRibbonButton
        Me.grpBoxes_fillWithExampleStyles = Me.Factory.CreateRibbonButton
        Me.mnuCloseDocuments1 = Me.Factory.CreateRibbonMenu
        Me.grpBoxes_ToES = Me.Factory.CreateRibbonButton
        Me.grpBoxes_ToBox1 = Me.Factory.CreateRibbonButton
        Me.grpBoxes_ToApp = Me.Factory.CreateRibbonButton
        Me.Separator36 = Me.Factory.CreateRibbonSeparator
        Me.grpBoxes_ToLT = Me.Factory.CreateRibbonButton
        Me.mnu_grpBoxes_Recommendations = Me.Factory.CreateRibbonMenu
        Me.grpBoxes_Recommendation = Me.Factory.CreateRibbonButton
        Me.grpBoxes_RecommendationES = Me.Factory.CreateRibbonButton
        Me.mnu_grpBoxes_Findings = Me.Factory.CreateRibbonMenu
        Me.grpBoxes_KeyFinding = Me.Factory.CreateRibbonButton
        Me.grpBoxes_KeyFindingES = Me.Factory.CreateRibbonButton
        Me.grpPullouts_mnu01 = Me.Factory.CreateRibbonMenu
        Me.grpPullouts_emphasisBox_Left = Me.Factory.CreateRibbonButton
        Me.grpPullouts_emphasisBox_Centre = Me.Factory.CreateRibbonButton
        Me.grpPullouts_emphasisBox_Right = Me.Factory.CreateRibbonButton
        Me.grpReport_mnu_CaseStudies = Me.Factory.CreateRibbonMenu
        Me.grpReport_mnu_CaseStudies_FullPage = Me.Factory.CreateRibbonButton
        Me.grpReport_mnu_CaseStudies_HalfPage = Me.Factory.CreateRibbonButton
        Me.Separator39 = Me.Factory.CreateRibbonSeparator
        Me.grpReport_mnu_CaseStudies_CaseStudyHeading = Me.Factory.CreateRibbonButton
        Me.Separator37 = Me.Factory.CreateRibbonSeparator
        Me.mnuCloseDocuments16 = Me.Factory.CreateRibbonMenu
        Me.grpFigures_Figure = Me.Factory.CreateRibbonButton
        Me.Separator44 = Me.Factory.CreateRibbonSeparator
        Me.grpFigures_Appendix = Me.Factory.CreateRibbonButton
        Me.Separator43 = Me.Factory.CreateRibbonSeparator
        Me.grpFigures_ES = Me.Factory.CreateRibbonButton
        Me.Separator42 = Me.Factory.CreateRibbonSeparator
        Me.grpFigures_LT = Me.Factory.CreateRibbonButton
        Me.Separator41 = Me.Factory.CreateRibbonSeparator
        Me.grpFigures_CaptionAndHeading = Me.Factory.CreateRibbonButton
        Me.grpFigures_CaptionAndHeadingApp = Me.Factory.CreateRibbonButton
        Me.grpFigures_CaptionAndHeadingES = Me.Factory.CreateRibbonButton
        Me.Separator40 = Me.Factory.CreateRibbonSeparator
        Me.grpFigures_StyleForSubHeadings = Me.Factory.CreateRibbonButton
        Me.mnuCloseDocuments33 = Me.Factory.CreateRibbonMenu
        Me.grpFigures_convertToES = Me.Factory.CreateRibbonButton
        Me.grpFigures_convertToStd = Me.Factory.CreateRibbonButton
        Me.grpFigures_convertToApp = Me.Factory.CreateRibbonButton
        Me.Separator45 = Me.Factory.CreateRibbonSeparator
        Me.grpFigures_convertToLT = Me.Factory.CreateRibbonButton
        Me.grpStylesRpt_mnu_tbls_00 = Me.Factory.CreateRibbonMenu
        Me.grpTbls_fillCellsWithCustomColour = Me.Factory.CreateRibbonButton
        Me.Separator46 = Me.Factory.CreateRibbonSeparator
        Me.grpTbls_setTableTextCustomColour = Me.Factory.CreateRibbonButton
        Me.Separator38 = Me.Factory.CreateRibbonSeparator
        Me.grpPlh_btn_buildCustomTable = Me.Factory.CreateRibbonButton
        Me.grpTbl_Styles = Me.Factory.CreateRibbonMenu
        Me.grpTbls_TableColumnHeadingsStyle = Me.Factory.CreateRibbonButton
        Me.grpTbls_TableUnitsRowStyle = Me.Factory.CreateRibbonButton
        Me.Separator47 = Me.Factory.CreateRibbonSeparator
        Me.grpTbls_Plh_mnu_TableListBulletsStyles = Me.Factory.CreateRibbonMenu
        Me.grpTbls_TableListBullet = Me.Factory.CreateRibbonButton
        Me.grpTbls_TableListBullet2 = Me.Factory.CreateRibbonButton
        Me.grpTbls_TableListBullet3 = Me.Factory.CreateRibbonButton
        Me.grpTbls_Plh_mnu_TableListNumberingStyles = Me.Factory.CreateRibbonMenu
        Me.grpTbls_ListNumber = Me.Factory.CreateRibbonButton
        Me.grpTbls_ListNumber2 = Me.Factory.CreateRibbonButton
        Me.grpTbls_ListNumber3 = Me.Factory.CreateRibbonButton
        Me.Separator48 = Me.Factory.CreateRibbonSeparator
        Me.grpTbls_Plh_mnu_SideHeadingStyles = Me.Factory.CreateRibbonMenu
        Me.grpTbls_TableSideHeading1 = Me.Factory.CreateRibbonButton
        Me.grpTbls_TableSideHeading2 = Me.Factory.CreateRibbonButton
        Me.grpTbls_Plh_mnu_QuoteStyles = Me.Factory.CreateRibbonMenu
        Me.grpTbls_Quote = Me.Factory.CreateRibbonButton
        Me.grpTbls_QuoteListBullet = Me.Factory.CreateRibbonButton
        Me.grpTbls_QuoteSource = Me.Factory.CreateRibbonButton
        Me.Separator49 = Me.Factory.CreateRibbonSeparator
        Me.grpTbl_Styles_ExampleStyleSets = Me.Factory.CreateRibbonMenu
        Me.grpTbls_StyleSet_TableQuote = Me.Factory.CreateRibbonButton
        Me.grpTbls_StyleSet_TableListBullets = Me.Factory.CreateRibbonButton
        Me.grpTbls_StyleSet_TableListNumbers = Me.Factory.CreateRibbonButton
        Me.Separator50 = Me.Factory.CreateRibbonSeparator
        Me.grpTbls_ColourCells = Me.Factory.CreateRibbonButton
        Me.grpTbls_ColourHeadingsRow = Me.Factory.CreateRibbonButton
        Me.grpTbls_ColourUnitsRow = Me.Factory.CreateRibbonButton
        Me.Separator51 = Me.Factory.CreateRibbonSeparator
        Me.grpTbls_AllBorders = Me.Factory.CreateRibbonButton
        Me.grpTbls_AllBordersRemove = Me.Factory.CreateRibbonButton
        Me.grpTbls_TableTextStyle = Me.Factory.CreateRibbonButton
        Me.mnuCloseDocuments4 = Me.Factory.CreateRibbonMenu
        Me.grpTbls_convertTabletoES = Me.Factory.CreateRibbonButton
        Me.grpTbls_convertTabletoStd = Me.Factory.CreateRibbonButton
        Me.grpTbls_convertTabletoApp = Me.Factory.CreateRibbonButton
        Me.Separator52 = Me.Factory.CreateRibbonSeparator
        Me.grpTbls_convertTabletoLT = Me.Factory.CreateRibbonButton
        Me.grpTbls_AllStyles_small = Me.Factory.CreateRibbonMenu
        Me.grpTbls_TableColumnHeadingsStyle_small = Me.Factory.CreateRibbonButton
        Me.grpTbls_TableUnitsRowStyle_small = Me.Factory.CreateRibbonButton
        Me.Separator57 = Me.Factory.CreateRibbonSeparator
        Me.grpTbls_Plh_mnu_TableListBulletsStyles_small = Me.Factory.CreateRibbonMenu
        Me.grpTbls_TableListBullet_small = Me.Factory.CreateRibbonButton
        Me.grpTbls_TableListBullet2_small = Me.Factory.CreateRibbonButton
        Me.grpTbls_TableListBullet3_small = Me.Factory.CreateRibbonButton
        Me.grpTbls_Plh_mnu_TableListNumberingStyles_small = Me.Factory.CreateRibbonMenu
        Me.grpTbls_ListNumber_small = Me.Factory.CreateRibbonButton
        Me.grpTbls_ListNumber2_small = Me.Factory.CreateRibbonButton
        Me.grpTbls_ListNumber3_small = Me.Factory.CreateRibbonButton
        Me.Separator58 = Me.Factory.CreateRibbonSeparator
        Me.grpPlh_mnu_TblSideHeadings_small = Me.Factory.CreateRibbonMenu
        Me.grpTbls_TableSideHeading1_small = Me.Factory.CreateRibbonButton
        Me.grpTbls_TableSideHeading2_small = Me.Factory.CreateRibbonButton
        Me.grpPlh_mnu_TblQuoteStyles_small = Me.Factory.CreateRibbonMenu
        Me.grpTbls_Quote_small = Me.Factory.CreateRibbonButton
        Me.grpTbls_QuoteListBullet_small = Me.Factory.CreateRibbonButton
        Me.grpTbls_QuoteSource_small = Me.Factory.CreateRibbonButton
        Me.Separator59 = Me.Factory.CreateRibbonSeparator
        Me.grpTbl_Styles_ExampleStyleSets_Small = Me.Factory.CreateRibbonMenu
        Me.grpTbls_StyleSet_TableQuote_small = Me.Factory.CreateRibbonButton
        Me.grpTbls_StyleSet_TableListBullets_small = Me.Factory.CreateRibbonButton
        Me.grpTbls_StyleSet_TableListNumbers_small = Me.Factory.CreateRibbonButton
        Me.grpTbls_TableTextStyle_small = Me.Factory.CreateRibbonButton
        Me.grpPlh_mnu_TblPlaceholders = Me.Factory.CreateRibbonMenu
        Me.grpTblsPlh_HeadingAndSource = Me.Factory.CreateRibbonButton
        Me.grpTblsPlh_HeadingAndSourceApp = Me.Factory.CreateRibbonButton
        Me.grpTblsPlh_HeadingAndSourceES = Me.Factory.CreateRibbonButton
        Me.Separator64 = Me.Factory.CreateRibbonSeparator
        Me.grpTblsPlh_CaptionAndHeading = Me.Factory.CreateRibbonButton
        Me.grpTblsPlh_CaptionAndHeadingApp = Me.Factory.CreateRibbonButton
        Me.grpTblsPlh_CaptionAndHeadingES = Me.Factory.CreateRibbonButton
        Me.Separator65 = Me.Factory.CreateRibbonSeparator
        Me.grpTblsPlh_AddTable_Simple = Me.Factory.CreateRibbonButton
        Me.grpPlh_mnu_SourceAndNote = Me.Factory.CreateRibbonMenu
        Me.grpTblsPlh_SourceLabelAndStyle = Me.Factory.CreateRibbonButton
        Me.grpTblsPlh_NoteLabelAndStyle = Me.Factory.CreateRibbonButton
        Me.Separator66 = Me.Factory.CreateRibbonSeparator
        Me.grpTblsPlh_SourceForOverType = Me.Factory.CreateRibbonButton
        Me.grpPlh_mnu_DeleteTable = Me.Factory.CreateRibbonMenu
        Me.grpTblsPlh_DeleteTable_fast = Me.Factory.CreateRibbonButton
        Me.grpTblsPlh_DeleteTable = Me.Factory.CreateRibbonButton
        Me.grp_special_AATableFormatting = Me.Factory.CreateRibbonGroup
        Me.tbPlh_mnu_convertPlhToHalfPage = Me.Factory.CreateRibbonMenu
        Me.grpTbl_mnu_AAPlh_To_HalfPlh_Left = Me.Factory.CreateRibbonButton
        Me.grpTbl_mnu_AAPlh_To_HalfPlh_Right = Me.Factory.CreateRibbonButton
        Me.Separator67 = Me.Factory.CreateRibbonSeparator
        Me.grpTbl_mnu_AAPlh_Reset_to_FullColumn = Me.Factory.CreateRibbonButton
        Me.Separator69 = Me.Factory.CreateRibbonSeparator
        Me.tbPlh_mnu_rapidFormat = Me.Factory.CreateRibbonMenu
        Me.grpTblsPlh_rapidFormat = Me.Factory.CreateRibbonButton
        Me.grpTblsPlh_rapidFormat_Encapsulated = Me.Factory.CreateRibbonButton
        Me.Separator68 = Me.Factory.CreateRibbonSeparator
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force = Me.Factory.CreateRibbonMenu
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_LT = Me.Factory.CreateRibbonButton
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_ES = Me.Factory.CreateRibbonButton
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_Body = Me.Factory.CreateRibbonButton
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_AP = Me.Factory.CreateRibbonButton
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force = Me.Factory.CreateRibbonMenu
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_LT = Me.Factory.CreateRibbonButton
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_ES = Me.Factory.CreateRibbonButton
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_Body = Me.Factory.CreateRibbonButton
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_AP = Me.Factory.CreateRibbonButton
        Me.Separator70 = Me.Factory.CreateRibbonSeparator
        Me.grpAATbls_mnu_editColumns = Me.Factory.CreateRibbonMenu
        Me.grpTblsEdit_InsertColumnRight = Me.Factory.CreateRibbonButton
        Me.grpTblsEdit_InsertColumnLeft = Me.Factory.CreateRibbonButton
        Me.Separator72 = Me.Factory.CreateRibbonSeparator
        Me.grpTblsEdit_Delete_Column = Me.Factory.CreateRibbonButton
        Me.grpAATbls_mnu_editRows = Me.Factory.CreateRibbonMenu
        Me.grpTblsEdit_InsertRowAbove = Me.Factory.CreateRibbonButton
        Me.grpTblsEdit_InsertRowBelow = Me.Factory.CreateRibbonButton
        Me.Separator73 = Me.Factory.CreateRibbonSeparator
        Me.grpTblsEdit_Delete_Row = Me.Factory.CreateRibbonButton
        Me.grpAATbls_mnu_AATableactions = Me.Factory.CreateRibbonMenu
        Me.grpTblsEdit_CopyTable = Me.Factory.CreateRibbonButton
        Me.Separator74 = Me.Factory.CreateRibbonSeparator
        Me.grpTblsEdit_PastePriorTable = Me.Factory.CreateRibbonButton
        Me.Separator75 = Me.Factory.CreateRibbonSeparator
        Me.grpTblsEdit_UndoTableAction = Me.Factory.CreateRibbonButton
        Me.Separator71 = Me.Factory.CreateRibbonSeparator
        Me.grp_Plh_TableColumns_mnu_more = Me.Factory.CreateRibbonMenu
        Me.grpTblsEdit_Convert_EncapsToStd = Me.Factory.CreateRibbonButton
        Me.grpTblsEdit_Convert_StdToEncaps = Me.Factory.CreateRibbonButton
        Me.Separator76 = Me.Factory.CreateRibbonSeparator
        Me.grpTblsEdit_Split_Table = Me.Factory.CreateRibbonButton
        Me.grp_floatingPlaceholders = Me.Factory.CreateRibbonGroup
        Me.grpReport_PlH_Handling = Me.Factory.CreateRibbonMenu
        Me.grpReport_PlH_LockToTop = Me.Factory.CreateRibbonButton
        Me.grpReport_PlHFloat_lock_toMarginsLeftAndTop = Me.Factory.CreateRibbonButton
        Me.grpReport_PlH_LockToParagraph = Me.Factory.CreateRibbonButton
        Me.grpReport_PlH_LockToParagraphAndColumn = Me.Factory.CreateRibbonButton
        Me.Separator77 = Me.Factory.CreateRibbonSeparator
        Me.grpReport_PlH_FloatEdgeToEdge = Me.Factory.CreateRibbonButton
        Me.grpReport_PlH_FloatWide = Me.Factory.CreateRibbonButton
        Me.grpReport_PlH_FloatMarginToMargin = Me.Factory.CreateRibbonButton
        Me.grpReport_PlH_ColumnWidth = Me.Factory.CreateRibbonButton
        Me.grpReport_PlH_TwoColumnWidth = Me.Factory.CreateRibbonButton
        Me.grpReport_PlH_convertToInline_findAllFloatingTables_2 = Me.Factory.CreateRibbonButton
        Me.grpReport_PlH_convertToInline = Me.Factory.CreateRibbonButton
        Me.grp_Plh_miscPlaceholders = Me.Factory.CreateRibbonGroup
        Me.grpPicts_PasteAsPic = Me.Factory.CreateRibbonButton
        Me.grpEquations_Numbered = Me.Factory.CreateRibbonButton
        Me.tab_aa_PagesAndSections = Me.Factory.CreateRibbonTab
        Me.grpRpt_CoversAndTOC = Me.Factory.CreateRibbonGroup
        Me.gal_CoverPages = Me.Factory.CreateRibbonGallery
        Me.gal_CoverPages_btn_deleteCoverPage = Me.Factory.CreateRibbonButton
        Me.grpCntsPages = Me.Factory.CreateRibbonMenu
        Me.grpContactsPages_FrontPage_AckOfCountry = Me.Factory.CreateRibbonButton
        Me.grpContactsPages_FrontPage = Me.Factory.CreateRibbonButton
        Me.Separator22 = Me.Factory.CreateRibbonSeparator
        Me.grpContactsPages_BackPage = Me.Factory.CreateRibbonButton
        Me.Separator23 = Me.Factory.CreateRibbonSeparator
        Me.grpContactsPages_mnu_2 = Me.Factory.CreateRibbonMenu
        Me.grpContactsPages_ReportTo = Me.Factory.CreateRibbonButton
        Me.grpContactsPages_ProposalTo = Me.Factory.CreateRibbonButton
        Me.Separator24 = Me.Factory.CreateRibbonSeparator
        Me.grpContactsPages_CopyrightStatement = Me.Factory.CreateRibbonButton
        Me.grpContactsPages_Disclaimer = Me.Factory.CreateRibbonButton
        Me.grpCoversToc_mnu2 = Me.Factory.CreateRibbonMenu
        Me.grpToc_TOC_insertSection = Me.Factory.CreateRibbonButton
        Me.Separator20 = Me.Factory.CreateRibbonSeparator
        Me.grpToc_TOC_insertLevels_1_to_1 = Me.Factory.CreateRibbonButton
        Me.grpToc_TOC_insertLevels_1_to_2 = Me.Factory.CreateRibbonButton
        Me.grpToc_TOC_insertLevels_1_to_3 = Me.Factory.CreateRibbonButton
        Me.Separator21 = Me.Factory.CreateRibbonSeparator
        Me.grpToc_TOC_update = Me.Factory.CreateRibbonButton
        Me.grpRpt_ImagePanels = Me.Factory.CreateRibbonGroup
        Me.grpCoversToc_mnu_Images = Me.Factory.CreateRibbonMenu
        Me.grpCpImages_ImageFromFile = Me.Factory.CreateRibbonButton
        Me.grpCpImages_ImageFromClip = Me.Factory.CreateRibbonButton
        Me.Separator26 = Me.Factory.CreateRibbonSeparator
        Me.grpCpImages_BackPanelFill_RawImageFromFile = Me.Factory.CreateRibbonButton
        Me.Separator25 = Me.Factory.CreateRibbonSeparator
        Me.grpCpImages_Delete_SmallPict = Me.Factory.CreateRibbonButton
        Me.grpImageHandling_mnu_ImgSection = Me.Factory.CreateRibbonMenu
        Me.grpImageHandling_insert_BackPanel = Me.Factory.CreateRibbonButton
        Me.Separator27 = Me.Factory.CreateRibbonSeparator
        Me.grpImageHandling_delete_BackPanel = Me.Factory.CreateRibbonButton
        Me.grpImageHandling_mnu_FillBackPanel = Me.Factory.CreateRibbonMenu
        Me.grpImageHandling_BackPanelFill_FromFile = Me.Factory.CreateRibbonButton
        Me.grpImageHandling_BackPanelFill_FromClipBoard = Me.Factory.CreateRibbonButton
        Me.Separator28 = Me.Factory.CreateRibbonSeparator
        Me.grpImageHandling_BackPanelFill_RawImageFromFile = Me.Factory.CreateRibbonButton
        Me.Separator29 = Me.Factory.CreateRibbonSeparator
        Me.grpImageHandling_Reset_backcolour = Me.Factory.CreateRibbonButton
        Me.grpImageHandling_Reset_backcolour_to_CaseStudyGrey = Me.Factory.CreateRibbonButton
        Me.grpImageHandling_Custom_backcolour = Me.Factory.CreateRibbonButton
        Me.Separator30 = Me.Factory.CreateRibbonSeparator
        Me.grpImageHandling_submnu_FillBackPanel_SetTransparency = Me.Factory.CreateRibbonMenu
        Me.submnu_SetTransparency_to_0 = Me.Factory.CreateRibbonButton
        Me.submnu_SetTransparency_to_25 = Me.Factory.CreateRibbonButton
        Me.submnu_SetTransparency_to_50 = Me.Factory.CreateRibbonButton
        Me.submnu_SetTransparency_to_75 = Me.Factory.CreateRibbonButton
        Me.submnu_SetTransparency_to_100 = Me.Factory.CreateRibbonButton
        Me.mnu_SetBackPanel_to_BannerHeight = Me.Factory.CreateRibbonButton
        Me.grpRpt_Report = Me.Factory.CreateRibbonGroup
        Me.grpRpt_btn_GlossaryAndAbbreviations_bblk = Me.Factory.CreateRibbonButton
        Me.grpReport_btn_newDivider_Chpt_bblk = Me.Factory.CreateRibbonButton
        Me.grpRpt_mnu_CreateExecSummary = Me.Factory.CreateRibbonMenu
        Me.grpExecSum_ExecSum_bblk = Me.Factory.CreateRibbonButton
        Me.grpExecSum_ExecSum_Grey_bblk = Me.Factory.CreateRibbonButton
        Me.Separator2 = Me.Factory.CreateRibbonSeparator
        Me.grpRpt_mnu_CreateRpt = Me.Factory.CreateRibbonMenu
        Me.grpReport_btn_buildPortraitReport = Me.Factory.CreateRibbonButton
        Me.Separator3 = Me.Factory.CreateRibbonSeparator
        Me.grpReprt_btn_buildLandscapeReport = Me.Factory.CreateRibbonButton
        Me.Separator4 = Me.Factory.CreateRibbonSeparator
        Me.grpReport_btn_buildAABrief = Me.Factory.CreateRibbonButton
        Me.grpRpt_mnu_NewChapter = Me.Factory.CreateRibbonMenu
        Me.grpRpt_mnu_btn_NewChapter_inFront_bblk = Me.Factory.CreateRibbonButton
        Me.grpRpt_mnu_btn_NewChapter_behind_bblk = Me.Factory.CreateRibbonButton
        Me.grpRpt_mnu_Bibliography = Me.Factory.CreateRibbonMenu
        Me.grpOther_bibliography_bblk = Me.Factory.CreateRibbonButton
        Me.grpOther_references_bblk = Me.Factory.CreateRibbonButton
        Me.grpOther_worksCited_bblk = Me.Factory.CreateRibbonButton
        Me.grpReport_btn_ToggleView = Me.Factory.CreateRibbonButton
        Me.grpRpt_mnu_RefreshDocument = Me.Factory.CreateRibbonMenu
        Me.grpViewTools_Refresh_Stationery_Ref = Me.Factory.CreateRibbonButton
        Me.Separator9 = Me.Factory.CreateRibbonSeparator
        Me.grpViewTools_Refresh_mnu_TOC = Me.Factory.CreateRibbonButton
        Me.grpViewTools_Refresh_mnu_Chapters = Me.Factory.CreateRibbonButton
        Me.grpViewTools_Refresh_mnu_Parts = Me.Factory.CreateRibbonButton
        Me.Separator10 = Me.Factory.CreateRibbonSeparator
        Me.grpViewTools_Refresh_mnu_Tables = Me.Factory.CreateRibbonButton
        Me.grpViewTools_Refresh_mnu_Figures = Me.Factory.CreateRibbonButton
        Me.grpViewTools_Refresh_mnu_Boxes = Me.Factory.CreateRibbonButton
        Me.Separator11 = Me.Factory.CreateRibbonSeparator
        Me.grpViewTools_Refresh_mnu_All = Me.Factory.CreateRibbonButton
        Me.Separator12 = Me.Factory.CreateRibbonSeparator
        Me.grpViewTools_Refresh_mnu_Every = Me.Factory.CreateRibbonButton
        Me.Separator13 = Me.Factory.CreateRibbonSeparator
        Me.mnu_grpViewTools_Refresh_btn_setRefFldNotBold = Me.Factory.CreateRibbonButton
        Me.grpRpt_mnu_ApplyColour = Me.Factory.CreateRibbonMenu
        Me.grpReport_mnu01_SelectedText = Me.Factory.CreateRibbonButton
        Me.grpReport_mnu01_SelectedTblCells = Me.Factory.CreateRibbonButton
        Me.grpReport_mnu01_ImageBackPanel = Me.Factory.CreateRibbonButton
        Me.Separator7 = Me.Factory.CreateRibbonSeparator
        Me.grpReport_mnu_CaseStudies_RecolourLogo_toWhite = Me.Factory.CreateRibbonButton
        Me.grpReport_mnu_CaseStudies_RecolourLogo_Reset = Me.Factory.CreateRibbonButton
        Me.Separator8 = Me.Factory.CreateRibbonSeparator
        Me.grpReport_mnu_CaseStudies_RecolourFooter_toWhite = Me.Factory.CreateRibbonButton
        Me.grpReport_mnu_CaseStudies_RecolourFooter_Reset = Me.Factory.CreateRibbonButton
        Me.grpRpt_Appendix = Me.Factory.CreateRibbonGroup
        Me.grpAppendix_mnu01 = Me.Factory.CreateRibbonMenu
        Me.grpAppendix_newAppPart = Me.Factory.CreateRibbonButton
        Me.grpAppendix_newAttPart = Me.Factory.CreateRibbonButton
        Me.grpReport_mnu_NewAppAtt = Me.Factory.CreateRibbonMenu
        Me.grpAppendix_newAppChapter_inFront_bblk = Me.Factory.CreateRibbonButton
        Me.grpAppendix_newAppChapter_behind_bblk = Me.Factory.CreateRibbonButton
        Me.grpRpt_sectOptions = Me.Factory.CreateRibbonGroup
        Me.mnuCloseDocuments000 = Me.Factory.CreateRibbonMenu
        Me.grpSectOptions_submnu_LndWidthOptions = Me.Factory.CreateRibbonMenu
        Me.grpSectOptions_sect_InsertSectionBounded_Lnd = Me.Factory.CreateRibbonButton
        Me.grpSectOptions_sect_InsertSectionBounded_Lnd_wide = Me.Factory.CreateRibbonButton
        Me.Menu3 = Me.Factory.CreateRibbonMenu
        Me.grpSectOptions_sect_InsertSectionBounded_Prt = Me.Factory.CreateRibbonButton
        Me.grpSectOptions_sect_InsertSectionBounded_Prt_wide = Me.Factory.CreateRibbonButton
        Me.Separator14 = Me.Factory.CreateRibbonSeparator
        Me.grpSectOptions_sect_InsertSection_AtSelection = Me.Factory.CreateRibbonButton
        Me.grpRpt_sectOptions_btn_delSection = Me.Factory.CreateRibbonButton
        Me.grpOther_mnuHFS = Me.Factory.CreateRibbonMenu
        Me.grpSectOptions_header_ClearTextandShapes = Me.Factory.CreateRibbonButton
        Me.grpSectOptions_footer_ClearText = Me.Factory.CreateRibbonButton
        Me.grpSectOptions_footer_ClearTextandPageNum = Me.Factory.CreateRibbonButton
        Me.Separator15 = Me.Factory.CreateRibbonSeparator
        Me.grpSectOptions_footer_clearSubTitleField = Me.Factory.CreateRibbonButton
        Me.Separator16 = Me.Factory.CreateRibbonSeparator
        Me.grpSectOptions_footer_resetText = Me.Factory.CreateRibbonButton
        Me.Separator17 = Me.Factory.CreateRibbonSeparator
        Me.grpOther_mnuHFS_sub00_restoreHF = Me.Factory.CreateRibbonMenu
        Me.grpSectOptions_hfs_restoreHF_ES = Me.Factory.CreateRibbonButton
        Me.grpSectOptions_hfs_restoreHF_RP = Me.Factory.CreateRibbonButton
        Me.grpSectOptions_hfs_restoreHF_AP = Me.Factory.CreateRibbonButton
        Me.grpSectOptions_mnu_ResetLndPrt = Me.Factory.CreateRibbonMenu
        Me.grpSectOptions_mnu_sub00_ResetLndPrt_Lnd = Me.Factory.CreateRibbonMenu
        Me.grpSectOptions_resetTo_Lnd_ES = Me.Factory.CreateRibbonButton
        Me.grpSectOptions_resetTo_Lnd_RP = Me.Factory.CreateRibbonButton
        Me.grpSectOptions_resetTo_Lnd_AP = Me.Factory.CreateRibbonButton
        Me.grpSectOptions_mnu_sub01_ResetLndPrt_Ln = Me.Factory.CreateRibbonMenu
        Me.grpSectOptions_resetTo_Prt_ES = Me.Factory.CreateRibbonButton
        Me.grpSectOptions_resetTo_Prt_RP = Me.Factory.CreateRibbonButton
        Me.grpSectOptions_resetTo_Prt_AP = Me.Factory.CreateRibbonButton
        Me.grpSectOptions_mnu_ResetResizeLandscape = Me.Factory.CreateRibbonMenu
        Me.grpSectOptions_resizeTo_Landscape = Me.Factory.CreateRibbonButton
        Me.grpSectOptions_resizeTo_Portrait = Me.Factory.CreateRibbonButton
        Me.Separator18 = Me.Factory.CreateRibbonSeparator
        Me.grpSectOptions_resize_toggleWidth = Me.Factory.CreateRibbonButton
        Me.mnu_grpReport_Columns = Me.Factory.CreateRibbonMenu
        Me.grpReport_Columns_04 = Me.Factory.CreateRibbonButton
        Me.grpReport_Columns_03 = Me.Factory.CreateRibbonButton
        Me.grpReport_Columns_02 = Me.Factory.CreateRibbonButton
        Me.grpReport_Columns_02_LeftWide = Me.Factory.CreateRibbonButton
        Me.grpReport_Columns_02_RightWide = Me.Factory.CreateRibbonButton
        Me.Separator19 = Me.Factory.CreateRibbonSeparator
        Me.grpReport_Columns_01 = Me.Factory.CreateRibbonButton
        Me.grpRpt_CoveringLetter = Me.Factory.CreateRibbonGroup
        Me.grpLetter_insertLetter = Me.Factory.CreateRibbonButton
        Me.grpLetter_insertMemo = Me.Factory.CreateRibbonButton
        Me.grpCoveringLetter_mnu6 = Me.Factory.CreateRibbonMenu
        Me.mnuCloseDocuments777 = Me.Factory.CreateRibbonMenu
        Me.grpLetter_Melbourne = Me.Factory.CreateRibbonButton
        Me.grpLetter_Sydney = Me.Factory.CreateRibbonButton
        Me.grpLetter_Brisbane = Me.Factory.CreateRibbonButton
        Me.grpLetter_Canberra = Me.Factory.CreateRibbonButton
        Me.grpLetter_Perth = Me.Factory.CreateRibbonButton
        Me.grpLetter_Adelaide = Me.Factory.CreateRibbonButton
        Me.Separator61 = Me.Factory.CreateRibbonSeparator
        Me.grpLetter_btn_forMemo = Me.Factory.CreateRibbonButton
        Me.mnuCloseDocuments11 = Me.Factory.CreateRibbonMenu
        Me.grpLetter_LtrHead1 = Me.Factory.CreateRibbonButton
        Me.grpLetter_LtrHead2 = Me.Factory.CreateRibbonButton
        Me.grpLetter_LtrHead3 = Me.Factory.CreateRibbonButton
        Me.grpLetter_delReport = Me.Factory.CreateRibbonButton
        Me.grp_WhatsNew = Me.Factory.CreateRibbonGroup
        Me.grpWhatsNew_Form = Me.Factory.CreateRibbonButton
        Me.grp_Fixes = Me.Factory.CreateRibbonGroup
        Me.grpFixes_Repairs = Me.Factory.CreateRibbonMenu
        Me.grpFixes_Repairs_remCharChar = Me.Factory.CreateRibbonButton
        Me.grpFixes_Repairs_remSpaces_indrCells = Me.Factory.CreateRibbonButton
        Me.grpFixes_Repairs_SetLanguage = Me.Factory.CreateRibbonButton
        Me.mnu_Pagination = Me.Factory.CreateRibbonMenu
        Me.grpFixes_RePaginate = Me.Factory.CreateRibbonButton
        Me.grpFixes_PaginateOff = Me.Factory.CreateRibbonButton
        Me.grpFixes_mnu_Other = Me.Factory.CreateRibbonMenu
        Me.mnu_Fixes_ScreenUpdating = Me.Factory.CreateRibbonMenu
        Me.grpFixes_ScreenUpdatingOff = Me.Factory.CreateRibbonButton
        Me.grpFixes_ScreenUpdatingOn = Me.Factory.CreateRibbonButton
        Me.tab_aa_Finalise = Me.Factory.CreateRibbonTab
        Me.grp_WaterMarks = Me.Factory.CreateRibbonGroup
        Me.grp_waterMark_mnu03 = Me.Factory.CreateRibbonMenu
        Me.grp_waterMark_cabinet_add = Me.Factory.CreateRibbonButton
        Me.grp_waterMark_commercial_add = Me.Factory.CreateRibbonButton
        Me.grp_waterMark_confidential_add = Me.Factory.CreateRibbonButton
        Me.grp_waterMark_restricted_add = Me.Factory.CreateRibbonButton
        Me.Separator84 = Me.Factory.CreateRibbonSeparator
        Me.grp_waterMark_atg_UNOFFICIAL_add = Me.Factory.CreateRibbonButton
        Me.grp_waterMark_atg_OFFICIAL_add = Me.Factory.CreateRibbonButton
        Me.grp_waterMark_atg_OFFICIAL_Sensitive_add = Me.Factory.CreateRibbonButton
        Me.Separator83 = Me.Factory.CreateRibbonSeparator
        Me.grp_waterMark_submnu01 = Me.Factory.CreateRibbonMenu
        Me.grp_waterMark_bold_sec = Me.Factory.CreateRibbonButton
        Me.grp_waterMark_NOTbold_sec = Me.Factory.CreateRibbonButton
        Me.grp_waterMark_colour_red_sec = Me.Factory.CreateRibbonButton
        Me.grp_waterMark_colour_grey_sec = Me.Factory.CreateRibbonButton
        Me.Separator82 = Me.Factory.CreateRibbonSeparator
        Me.grp_waterMark_alignment_Centre_sec = Me.Factory.CreateRibbonButton
        Me.grp_waterMark_alignment_Right_sec = Me.Factory.CreateRibbonButton
        Me.Separator81 = Me.Factory.CreateRibbonSeparator
        Me.grp_waterMark_forceSec_StyleToDefault = Me.Factory.CreateRibbonButton
        Me.grp_waterMark_mnu01 = Me.Factory.CreateRibbonMenu
        Me.grp_waterMark_removeAll = Me.Factory.CreateRibbonButton
        Me.Separator78 = Me.Factory.CreateRibbonSeparator
        Me.grp_waterMark_mnu04 = Me.Factory.CreateRibbonMenu
        Me.grp_waterMark_removeSec = Me.Factory.CreateRibbonButton
        Me.grp_waterMark_removeStat = Me.Factory.CreateRibbonButton
        Me.grp_waterMark_mnu05 = Me.Factory.CreateRibbonMenu
        Me.grp_waterMark_removeSec_fromSect = Me.Factory.CreateRibbonButton
        Me.grp_waterMark_removeStat_fromSect = Me.Factory.CreateRibbonButton
        Me.grp_waterMark_mnu02 = Me.Factory.CreateRibbonMenu
        Me.grp_waterMark_draft_add = Me.Factory.CreateRibbonButton
        Me.grp_waterMark_draftOnly_add = Me.Factory.CreateRibbonButton
        Me.Separator79 = Me.Factory.CreateRibbonSeparator
        Me.grp_waterMark_colour_red_stat = Me.Factory.CreateRibbonButton
        Me.grp_waterMark_colour_grey_stat = Me.Factory.CreateRibbonButton
        Me.Separator80 = Me.Factory.CreateRibbonSeparator
        Me.grp_waterMark_forceStat_StyleToDefault = Me.Factory.CreateRibbonButton
        Me.grp_PgNumMgmnt = Me.Factory.CreateRibbonGroup
        Me.tabFin_mnu_PageNumFormatting = Me.Factory.CreateRibbonMenu
        Me.grpFixes_ApplyEsNumbering = Me.Factory.CreateRibbonButton
        Me.grpFixes_ApplyStdNumbering = Me.Factory.CreateRibbonButton
        Me.grpFixes_ApplyAppNumbering = Me.Factory.CreateRibbonButton
        Me.Separator86 = Me.Factory.CreateRibbonSeparator
        Me.grpFixes_ContinueNumbering = Me.Factory.CreateRibbonButton
        Me.grpFixes_RestartNumbering = Me.Factory.CreateRibbonButton
        Me.Separator85 = Me.Factory.CreateRibbonSeparator
        Me.grpFixes_getNumberingDialog = Me.Factory.CreateRibbonButton
        Me.tabFin_mnu_PgNumMgmnt_ReNum = Me.Factory.CreateRibbonMenu
        Me.grp_PgNumMgmnt_ReNum_std = Me.Factory.CreateRibbonButton
        Me.grp_PgNumMgmnt_ReNum_2Part = Me.Factory.CreateRibbonButton
        Me.grp_Finalise = Me.Factory.CreateRibbonGroup
        Me.grp_Finalise_mnu01 = Me.Factory.CreateRibbonMenu
        Me.grp_Finalise_CrossRefError = Me.Factory.CreateRibbonButton
        Me.Separator88 = Me.Factory.CreateRibbonSeparator
        Me.grp_Finalise_DoAll = Me.Factory.CreateRibbonButton
        Me.grp_Finalise_AllFunctions = Me.Factory.CreateRibbonMenu
        Me.grp_Finalise_upDateCopyrightNotice = Me.Factory.CreateRibbonButton
        Me.Separator87 = Me.Factory.CreateRibbonSeparator
        Me.grp_Finalise_updateFields = Me.Factory.CreateRibbonButton
        Me.grp_Finalise_setFootersToBold = Me.Factory.CreateRibbonButton
        Me.grp_Finalise_RefreshTOC = Me.Factory.CreateRibbonButton
        Me.grp_Finalise_CrossRefFlds_setToArialNarrow = Me.Factory.CreateRibbonButton
        Me.grp_Finalise_CrossRefFlds_setRefFldNotBold = Me.Factory.CreateRibbonButton
        Me.grpFixes_Repairs_delSpace1_betweenWords = Me.Factory.CreateRibbonButton
        Me.grpFixes_Repairs_delSpace1_atSentenceEnd = Me.Factory.CreateRibbonButton
        Me.grpWCAG = Me.Factory.CreateRibbonGroup
        Me.tabFin_mnu_AccessibilityTools = Me.Factory.CreateRibbonMenu
        Me.grpWCAG_notesOnAccessibility = Me.Factory.CreateRibbonButton
        Me.grpWCAG_convertThisDoc = Me.Factory.CreateRibbonButton
        Me.Separator89 = Me.Factory.CreateRibbonSeparator
        Me.grpWCAG_mnu_ContrastControl = Me.Factory.CreateRibbonMenu
        Me.grpWCAG_mnu_SetTransparency = Me.Factory.CreateRibbonMenu
        Me.grpWCAG_mnu_SetTransparency_to_0 = Me.Factory.CreateRibbonButton
        Me.grpWCAG_mnu_SetTransparency_to_25 = Me.Factory.CreateRibbonButton
        Me.grpWCAG_mnu_SetTransparency_to_50 = Me.Factory.CreateRibbonButton
        Me.grpWCAG_mnu_SetTransparency_to_75 = Me.Factory.CreateRibbonButton
        Me.grpWCAG_mnu_SetTransparency_to_100 = Me.Factory.CreateRibbonButton
        Me.grpWCAG_tool_convertAllStyles_toBlack = Me.Factory.CreateRibbonButton
        Me.grpWCAG_tool_tableHeaderColour_all = Me.Factory.CreateRibbonButton
        Me.grpReport_PlH_convertToInline_findAllFloatingTables = Me.Factory.CreateRibbonButton
        Me.grpRbn_Mgmnt = Me.Factory.CreateRibbonGroup
        Me.grpRbn_Mgmnt_mnu_00 = Me.Factory.CreateRibbonMenu
        Me.grpRbn_Mgmnt_removeRbn = Me.Factory.CreateRibbonButton
        Me.Separator54 = Me.Factory.CreateRibbonSeparator
        Me.grpTst_LoadFromWeb = Me.Factory.CreateRibbonGroup
        Me.grpRbn_Downloads_mnu_00 = Me.Factory.CreateRibbonMenu
        Me.grpTst_LoadFromResources_getStylesGuide = Me.Factory.CreateRibbonButton
        Me.grpTst_LoadFromResources_getStylesGuide_Accessible = Me.Factory.CreateRibbonButton
        Me.grpTst_LoadFromWeb_getStylesGuide = Me.Factory.CreateRibbonButton
        Me.grpTst_LoadFromWeb_getStylesGuide_Accessible = Me.Factory.CreateRibbonButton
        Me.Separator53 = Me.Factory.CreateRibbonSeparator
        Me.grpTst_LoadFromResources_getTemplate = Me.Factory.CreateRibbonButton
        Me.grpTst_LoadFromResources_getThemeFile = Me.Factory.CreateRibbonButton
        Me.Separator56 = Me.Factory.CreateRibbonSeparator
        Me.grpTst_LoadFromResources_getRptPrtExample = Me.Factory.CreateRibbonButton
        Me.grpTst_LoadFromResources_getRptLndExample = Me.Factory.CreateRibbonButton
        Me.grpTst_LoadFromResources_getRptBrfExample = Me.Factory.CreateRibbonButton
        Me.grpTst_LoadFromWeb_getTemplate = Me.Factory.CreateRibbonButton
        Me.grpMetaData = Me.Factory.CreateRibbonGroup
        Me.grpMetaData_Remove_FromDoc = Me.Factory.CreateRibbonButton
        Me.grpTestTools = Me.Factory.CreateRibbonGroup
        Me.grpTest_pgNum_getTagStyleMap = Me.Factory.CreateRibbonButton
        Me.tab_aa_Home = Me.Factory.CreateRibbonTab
        Me.grp_AA_ThemeandHome = Me.Factory.CreateRibbonGroup
        Me.Menu1 = Me.Factory.CreateRibbonMenu
        Me.tabThms_mnu_Set_btn_applyAATheme = Me.Factory.CreateRibbonButton
        Me.tabThms_mnu_Set_btn__applyAAThemeStylesTemplate = Me.Factory.CreateRibbonButton
        Me.Separator6 = Me.Factory.CreateRibbonSeparator
        Me.tabThms_mnu_Set_btn_attachNormalTemplate = Me.Factory.CreateRibbonButton
        Me.tabThms_mnu_Set_btn_attachAATemplate = Me.Factory.CreateRibbonButton
        Me.Separator60 = Me.Factory.CreateRibbonSeparator
        Me.tabThms_mnu_Set_btn_getAttachedTemplate = Me.Factory.CreateRibbonButton
        Me.tabThms_mnu_Set_btn_ActivateTabPGS = Me.Factory.CreateRibbonButton
        Me.tabThms_mnu_Set_btn_PGSToggle = Me.Factory.CreateRibbonButton
        Me.tabThms_mnu_resetStyles1 = Me.Factory.CreateRibbonMenu
        Me.tabThms_btn_resetStylesForRptPrt = Me.Factory.CreateRibbonButton
        Me.tabThms_btn_resetStylesForRptLnd = Me.Factory.CreateRibbonButton
        Me.tabThms_btn_resetStylesForRptBrf = Me.Factory.CreateRibbonButton
        Me.btn_colorPicker = Me.Factory.CreateRibbonButton
        Me.grp_buildDocuments = Me.Factory.CreateRibbonGroup
        Me.tbHome_mnu_CreateReport = Me.Factory.CreateRibbonMenu
        Me.grpReport_tbHome_btn_buildPortraitReport = Me.Factory.CreateRibbonButton
        Me.Separator92 = Me.Factory.CreateRibbonSeparator
        Me.grpReport_tbHome_btn_buildLandscapeReport = Me.Factory.CreateRibbonButton
        Me.Separator93 = Me.Factory.CreateRibbonSeparator
        Me.grpReport_tbHome_btn_buildAABrief = Me.Factory.CreateRibbonButton
        Me.Separator90 = Me.Factory.CreateRibbonSeparator
        Me.tbHome_grpLetter_standaloneLetter = Me.Factory.CreateRibbonButton
        Me.tbHome_grpLetter_standaloneMemo = Me.Factory.CreateRibbonButton
        Me.tbHome_mnu_contactDetails = Me.Factory.CreateRibbonMenu
        Me.tbHome_mnu_contactDetails_letter = Me.Factory.CreateRibbonMenu
        Me.tbHome_mnu_contactDetails_letter_Melbourne = Me.Factory.CreateRibbonButton
        Me.tbHome_mnu_contactDetails_letter_Sydney = Me.Factory.CreateRibbonButton
        Me.tbHome_mnu_contactDetails_letter_Brisbane = Me.Factory.CreateRibbonButton
        Me.tbHome_mnu_contactDetails_letter_Canberra = Me.Factory.CreateRibbonButton
        Me.tbHome_mnu_contactDetails_letter_Perth = Me.Factory.CreateRibbonButton
        Me.tbHome_mnu_contactDetails_letter_Adelaide = Me.Factory.CreateRibbonButton
        Me.tbHome_mnu_contactDetails_memo = Me.Factory.CreateRibbonButton
        Me.Separator91 = Me.Factory.CreateRibbonSeparator
        Me.btn_update_Fields = Me.Factory.CreateRibbonButton
        Me.Separator95 = Me.Factory.CreateRibbonSeparator
        Me.tbHome_btn_ToggleView = Me.Factory.CreateRibbonButton
        Me.Separator94 = Me.Factory.CreateRibbonSeparator
        Me.tbHome_btn_Help = Me.Factory.CreateRibbonButton
        Me.grpTest = Me.Factory.CreateRibbonGroup
        Me.grpTest_btn_cloneDoc = Me.Factory.CreateRibbonButton
        Me.grpTest_btn_getTimeStamp = Me.Factory.CreateRibbonButton
        Me.grpSectOptions_sect_InsertSection_InFront = Me.Factory.CreateRibbonButton
        Me.grpSectOptions_sect_InsertSection_Behind = Me.Factory.CreateRibbonButton
        Me.grp_SwBuild = Me.Factory.CreateRibbonGroup
        Me.grpRpt_btn_GlossaryAndAbbreviations = Me.Factory.CreateRibbonButton
        Me.grpReport_btn_newDivider_Chpt = Me.Factory.CreateRibbonButton
        Me.Menu5 = Me.Factory.CreateRibbonMenu
        Me.grpExecSum_ExecSum = Me.Factory.CreateRibbonButton
        Me.grpExecSum_ExecSum_Grey = Me.Factory.CreateRibbonButton
        Me.Menu2df = Me.Factory.CreateRibbonMenu
        Me.tabPgs_grpRpt_btn_buildPrtReport_sw = Me.Factory.CreateRibbonButton
        Me.Separator62 = Me.Factory.CreateRibbonSeparator
        Me.tabPgs_grpRpt_btn_buildLndReport_sw = Me.Factory.CreateRibbonButton
        Me.Separator63 = Me.Factory.CreateRibbonSeparator
        Me.tabPgs_grpRpt_btn_buildBrfReport_sw = Me.Factory.CreateRibbonButton
        Me.Menu2 = Me.Factory.CreateRibbonMenu
        Me.grpRpt_mnu_btn_NewChapter_inFront = Me.Factory.CreateRibbonButton
        Me.grpRpt_mnu_btn_NewChapter_behind = Me.Factory.CreateRibbonButton
        Me.Menu6 = Me.Factory.CreateRibbonMenu
        Me.grpOther_bibliography = Me.Factory.CreateRibbonButton
        Me.grpOther_references = Me.Factory.CreateRibbonButton
        Me.grpOther_worksCited = Me.Factory.CreateRibbonButton
        Me.Menu4 = Me.Factory.CreateRibbonMenu
        Me.grpAppendix_newAppChapter_inFront = Me.Factory.CreateRibbonButton
        Me.grpAppendix_newAppChapter_behind = Me.Factory.CreateRibbonButton
        Me.grpLetters_mnu_swBuilds = Me.Factory.CreateRibbonMenu
        Me.grpLetter_insertLetter_swBuild = Me.Factory.CreateRibbonButton
        Me.grpLetter_insertMemo_swBuild = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.btn_ApplyStdTheme_Manually = Me.Factory.CreateRibbonButton
        Me.mnu_makeStyles = Me.Factory.CreateRibbonMenu
        Me.btn_styles_makeTableText = Me.Factory.CreateRibbonButton
        Me.tab_aa_Styles.SuspendLayout()
        Me.grp_Styles_AAThemes.SuspendLayout()
        Me.grpStyles_CoverPage.SuspendLayout()
        Me.grpStyles_Report.SuspendLayout()
        Me.grpStyles_NoNum.SuspendLayout()
        Me.grpStyles_Appendices.SuspendLayout()
        Me.grpStyles_Text.SuspendLayout()
        Me.grpStyles_Lists.SuspendLayout()
        Me.grpStyles_Emphasis.SuspendLayout()
        Me.grpStyles_resetStyles.SuspendLayout()
        Me.grpStyles_resetCaptions.SuspendLayout()
        Me.tab_aa_Placeholders.SuspendLayout()
        Me.grp_PlaceHolders.SuspendLayout()
        Me.grp_special_AATableFormatting.SuspendLayout()
        Me.grp_floatingPlaceholders.SuspendLayout()
        Me.grp_Plh_miscPlaceholders.SuspendLayout()
        Me.tab_aa_PagesAndSections.SuspendLayout()
        Me.grpRpt_CoversAndTOC.SuspendLayout()
        Me.grpRpt_ImagePanels.SuspendLayout()
        Me.grpRpt_Report.SuspendLayout()
        Me.grpRpt_Appendix.SuspendLayout()
        Me.grpRpt_sectOptions.SuspendLayout()
        Me.grpRpt_CoveringLetter.SuspendLayout()
        Me.grp_WhatsNew.SuspendLayout()
        Me.grp_Fixes.SuspendLayout()
        Me.tab_aa_Finalise.SuspendLayout()
        Me.grp_WaterMarks.SuspendLayout()
        Me.grp_PgNumMgmnt.SuspendLayout()
        Me.grp_Finalise.SuspendLayout()
        Me.grpWCAG.SuspendLayout()
        Me.grpRbn_Mgmnt.SuspendLayout()
        Me.grpTst_LoadFromWeb.SuspendLayout()
        Me.grpMetaData.SuspendLayout()
        Me.grpTestTools.SuspendLayout()
        Me.tab_aa_Home.SuspendLayout()
        Me.grp_AA_ThemeandHome.SuspendLayout()
        Me.grp_buildDocuments.SuspendLayout()
        Me.grpTest.SuspendLayout()
        Me.grp_SwBuild.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.SuspendLayout()
        '
        'tab_aa_Styles
        '
        Me.tab_aa_Styles.Groups.Add(Me.grp_Styles_AAThemes)
        Me.tab_aa_Styles.Groups.Add(Me.grpStyles_CoverPage)
        Me.tab_aa_Styles.Groups.Add(Me.grpStyles_Report)
        Me.tab_aa_Styles.Groups.Add(Me.grpStyles_NoNum)
        Me.tab_aa_Styles.Groups.Add(Me.grpStyles_Appendices)
        Me.tab_aa_Styles.Groups.Add(Me.grpStyles_Text)
        Me.tab_aa_Styles.Groups.Add(Me.grpStyles_Lists)
        Me.tab_aa_Styles.Groups.Add(Me.grpStyles_Emphasis)
        Me.tab_aa_Styles.Groups.Add(Me.grpStyles_resetStyles)
        Me.tab_aa_Styles.Groups.Add(Me.grpStyles_resetCaptions)
        Me.tab_aa_Styles.KeyTip = "JS"
        Me.tab_aa_Styles.Label = "Styles"
        Me.tab_aa_Styles.Name = "tab_aa_Styles"
        Me.tab_aa_Styles.Position = Me.Factory.RibbonPosition.AfterOfficeId("TabHome")
        '
        'grp_Styles_AAThemes
        '
        Me.grp_Styles_AAThemes.Items.Add(Me.mnu_SetTheme)
        Me.grp_Styles_AAThemes.Label = "ACIl Allen Theme"
        Me.grp_Styles_AAThemes.Name = "grp_Styles_AAThemes"
        Me.grp_Styles_AAThemes.Visible = False
        '
        'mnu_SetTheme
        '
        Me.mnu_SetTheme.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.mnu_SetTheme.Description = "Allows the user to apply standard ACIL Allen Theme"
        Me.mnu_SetTheme.Items.Add(Me.xbtn__mnuThemes_set_AATheme)
        Me.mnu_SetTheme.Items.Add(Me.xbtn__mnuThemes_set_AAThemeAndStyles)
        Me.mnu_SetTheme.Items.Add(Me.Separator1)
        Me.mnu_SetTheme.Items.Add(Me.xbtn_mnuThemes_ActivateTabPGS)
        Me.mnu_SetTheme.Items.Add(Me.xbtn_mnuThemes_PGSToggle)
        Me.mnu_SetTheme.Label = "Set  AA Theme"
        Me.mnu_SetTheme.Name = "mnu_SetTheme"
        Me.mnu_SetTheme.OfficeImageId = "ThemesGallery"
        Me.mnu_SetTheme.ShowImage = True
        '
        'xbtn__mnuThemes_set_AATheme
        '
        Me.xbtn__mnuThemes_set_AATheme.Label = "Apply the ACIL Allen theme to the current document"
        Me.xbtn__mnuThemes_set_AATheme.Name = "xbtn__mnuThemes_set_AATheme"
        Me.xbtn__mnuThemes_set_AATheme.OfficeImageId = "ThemesGallery"
        Me.xbtn__mnuThemes_set_AATheme.ShowImage = True
        '
        'xbtn__mnuThemes_set_AAThemeAndStyles
        '
        Me.xbtn__mnuThemes_set_AAThemeAndStyles.Label = "Apply ACIL Allen theme and styles to the current document"
        Me.xbtn__mnuThemes_set_AAThemeAndStyles.Name = "xbtn__mnuThemes_set_AAThemeAndStyles"
        Me.xbtn__mnuThemes_set_AAThemeAndStyles.OfficeImageId = "ThemesGallery"
        Me.xbtn__mnuThemes_set_AAThemeAndStyles.ShowImage = True
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'xbtn_mnuThemes_ActivateTabPGS
        '
        Me.xbtn_mnuThemes_ActivateTabPGS.Label = "Activate Pages and Sections tab"
        Me.xbtn_mnuThemes_ActivateTabPGS.Name = "xbtn_mnuThemes_ActivateTabPGS"
        Me.xbtn_mnuThemes_ActivateTabPGS.OfficeImageId = "ThemesGallery"
        Me.xbtn_mnuThemes_ActivateTabPGS.ShowImage = True
        '
        'xbtn_mnuThemes_PGSToggle
        '
        Me.xbtn_mnuThemes_PGSToggle.Label = "Make Pages and Sections tab invisible"
        Me.xbtn_mnuThemes_PGSToggle.Name = "xbtn_mnuThemes_PGSToggle"
        Me.xbtn_mnuThemes_PGSToggle.OfficeImageId = "ThemesGallery"
        Me.xbtn_mnuThemes_PGSToggle.ShowImage = True
        '
        'grpStyles_CoverPage
        '
        Me.grpStyles_CoverPage.Items.Add(Me.grpStylesES_StyleSet)
        Me.grpStyles_CoverPage.Items.Add(Me.grpStylesES_Heading2_ES)
        Me.grpStyles_CoverPage.Items.Add(Me.grpStylesES_Heading3_ES)
        Me.grpStyles_CoverPage.Items.Add(Me.grpStylesES_Heading4_ES)
        Me.grpStyles_CoverPage.Items.Add(Me.grpStylesES_Heading5_ES)
        Me.grpStyles_CoverPage.Label = "Exec Sum"
        Me.grpStyles_CoverPage.Name = "grpStyles_CoverPage"
        '
        'grpStylesES_StyleSet
        '
        Me.grpStylesES_StyleSet.KeyTip = "ES"
        Me.grpStylesES_StyleSet.Label = "ES Style Set"
        Me.grpStylesES_StyleSet.Name = "grpStylesES_StyleSet"
        Me.grpStylesES_StyleSet.OfficeImageId = "FunctionsLogicalInsertGallery"
        Me.grpStylesES_StyleSet.ShowImage = True
        '
        'grpStylesES_Heading2_ES
        '
        Me.grpStylesES_Heading2_ES.KeyTip = "E2"
        Me.grpStylesES_Heading2_ES.Label = "ES H2"
        Me.grpStylesES_Heading2_ES.Name = "grpStylesES_Heading2_ES"
        Me.grpStylesES_Heading2_ES.OfficeImageId = "BevelShapeGallery"
        Me.grpStylesES_Heading2_ES.ShowImage = True
        Me.grpStylesES_Heading2_ES.SuperTip = """ES Heading 2 style - only use in the Executive Summary"""
        '
        'grpStylesES_Heading3_ES
        '
        Me.grpStylesES_Heading3_ES.KeyTip = "E3"
        Me.grpStylesES_Heading3_ES.Label = "ES H3"
        Me.grpStylesES_Heading3_ES.Name = "grpStylesES_Heading3_ES"
        Me.grpStylesES_Heading3_ES.OfficeImageId = "BevelShapeGallery"
        Me.grpStylesES_Heading3_ES.ShowImage = True
        Me.grpStylesES_Heading3_ES.SuperTip = """ES Heading 3 style - only use in the Executive Summary"""
        '
        'grpStylesES_Heading4_ES
        '
        Me.grpStylesES_Heading4_ES.KeyTip = "E4"
        Me.grpStylesES_Heading4_ES.Label = "ES H4"
        Me.grpStylesES_Heading4_ES.Name = "grpStylesES_Heading4_ES"
        Me.grpStylesES_Heading4_ES.OfficeImageId = "BevelShapeGallery"
        Me.grpStylesES_Heading4_ES.ShowImage = True
        Me.grpStylesES_Heading4_ES.SuperTip = """ES Heading 4 style - only use in the Executive Summary"""
        '
        'grpStylesES_Heading5_ES
        '
        Me.grpStylesES_Heading5_ES.KeyTip = "E5"
        Me.grpStylesES_Heading5_ES.Label = "ES H5"
        Me.grpStylesES_Heading5_ES.Name = "grpStylesES_Heading5_ES"
        Me.grpStylesES_Heading5_ES.OfficeImageId = "BevelShapeGallery"
        Me.grpStylesES_Heading5_ES.ShowImage = True
        Me.grpStylesES_Heading5_ES.SuperTip = """ES Heading 5 style - only use in the Executive Summary"""
        '
        'grpStyles_Report
        '
        Me.grpStyles_Report.Items.Add(Me.grpStylesRpt_StyleSet)
        Me.grpStyles_Report.Items.Add(Me.grpStylesRpt_Heading2_Rpt)
        Me.grpStyles_Report.Items.Add(Me.grpStylesRpt_Heading3_Rpt)
        Me.grpStyles_Report.Items.Add(Me.grpStylesRpt_Heading4_Rpt)
        Me.grpStyles_Report.Items.Add(Me.grpStylesRpt_Heading5_Rpt)
        Me.grpStyles_Report.Items.Add(Me.Separator5)
        Me.grpStyles_Report.Items.Add(Me.grpStyles_mnu_Heading3Numbering)
        Me.grpStyles_Report.Label = "Report"
        Me.grpStyles_Report.Name = "grpStyles_Report"
        '
        'grpStylesRpt_StyleSet
        '
        Me.grpStylesRpt_StyleSet.KeyTip = "RS"
        Me.grpStylesRpt_StyleSet.Label = "Report StyleSet"
        Me.grpStylesRpt_StyleSet.Name = "grpStylesRpt_StyleSet"
        Me.grpStylesRpt_StyleSet.OfficeImageId = "FunctionsLogicalInsertGallery"
        Me.grpStylesRpt_StyleSet.ShowImage = True
        Me.grpStylesRpt_StyleSet.SuperTip = """This button will insert a report skeleton that uses most of the major report sty" &
    "les"""
        '
        'grpStylesRpt_Heading2_Rpt
        '
        Me.grpStylesRpt_Heading2_Rpt.KeyTip = "R2"
        Me.grpStylesRpt_Heading2_Rpt.Label = "Heading 2"
        Me.grpStylesRpt_Heading2_Rpt.Name = "grpStylesRpt_Heading2_Rpt"
        Me.grpStylesRpt_Heading2_Rpt.OfficeImageId = "BevelShapeGallery"
        Me.grpStylesRpt_Heading2_Rpt.ShowImage = True
        Me.grpStylesRpt_Heading2_Rpt.SuperTip = """Heading 2  style - use in the 'Chapter' parts of the report"""
        '
        'grpStylesRpt_Heading3_Rpt
        '
        Me.grpStylesRpt_Heading3_Rpt.KeyTip = "R3"
        Me.grpStylesRpt_Heading3_Rpt.Label = "Heading 3"
        Me.grpStylesRpt_Heading3_Rpt.Name = "grpStylesRpt_Heading3_Rpt"
        Me.grpStylesRpt_Heading3_Rpt.OfficeImageId = "BevelShapeGallery"
        Me.grpStylesRpt_Heading3_Rpt.ShowImage = True
        Me.grpStylesRpt_Heading3_Rpt.SuperTip = " ""Heading 3  style - use in the 'Chapter' parts of the report"""
        '
        'grpStylesRpt_Heading4_Rpt
        '
        Me.grpStylesRpt_Heading4_Rpt.KeyTip = "R4"
        Me.grpStylesRpt_Heading4_Rpt.Label = "Heading 4"
        Me.grpStylesRpt_Heading4_Rpt.Name = "grpStylesRpt_Heading4_Rpt"
        Me.grpStylesRpt_Heading4_Rpt.OfficeImageId = "BevelShapeGallery"
        Me.grpStylesRpt_Heading4_Rpt.ShowImage = True
        Me.grpStylesRpt_Heading4_Rpt.SuperTip = " ""Heading 4  style - use in the 'Chapter' parts of the report"""
        '
        'grpStylesRpt_Heading5_Rpt
        '
        Me.grpStylesRpt_Heading5_Rpt.KeyTip = "R5"
        Me.grpStylesRpt_Heading5_Rpt.Label = "Heading 5"
        Me.grpStylesRpt_Heading5_Rpt.Name = "grpStylesRpt_Heading5_Rpt"
        Me.grpStylesRpt_Heading5_Rpt.OfficeImageId = "BevelShapeGallery"
        Me.grpStylesRpt_Heading5_Rpt.ShowImage = True
        Me.grpStylesRpt_Heading5_Rpt.SuperTip = """Heading 5  style - use in the 'Chapter' parts of the report"""
        '
        'Separator5
        '
        Me.Separator5.Name = "Separator5"
        '
        'grpStyles_mnu_Heading3Numbering
        '
        Me.grpStyles_mnu_Heading3Numbering.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grpStyles_mnu_Heading3Numbering.Items.Add(Me.grpStyles_mnu_Heading3Numbering_btn_on)
        Me.grpStyles_mnu_Heading3Numbering.Items.Add(Me.grpStyles_mnu_Heading3Numbering_btn_off)
        Me.grpStyles_mnu_Heading3Numbering.KeyTip = "HN"
        Me.grpStyles_mnu_Heading3Numbering.Label = "Heading 3 Numbering"
        Me.grpStyles_mnu_Heading3Numbering.Name = "grpStyles_mnu_Heading3Numbering"
        Me.grpStyles_mnu_Heading3Numbering.OfficeImageId = "BevelShapeGallery"
        Me.grpStyles_mnu_Heading3Numbering.ShowImage = True
        Me.grpStyles_mnu_Heading3Numbering.SuperTip = """These menu items allow you to turn on/off numbered headings for Heading level 3 " &
    "(Body and AP).&#13;&#13;The default state for Heading level 3 (Body and AP) is '" &
    "no number'."""
        '
        'grpStyles_mnu_Heading3Numbering_btn_on
        '
        Me.grpStyles_mnu_Heading3Numbering_btn_on.Label = "Heading 3 numbering - O&n"
        Me.grpStyles_mnu_Heading3Numbering_btn_on.Name = "grpStyles_mnu_Heading3Numbering_btn_on"
        Me.grpStyles_mnu_Heading3Numbering_btn_on.OfficeImageId = "BevelShapeGallery"
        Me.grpStyles_mnu_Heading3Numbering_btn_on.ShowImage = True
        Me.grpStyles_mnu_Heading3Numbering_btn_on.SuperTip = """This button will adjust the body and appendix heading 3 levels to a numbered opt" &
    "ion (e.g. A.B.C)"""
        '
        'grpStyles_mnu_Heading3Numbering_btn_off
        '
        Me.grpStyles_mnu_Heading3Numbering_btn_off.Label = "Heading 3 numbering - O&ff"
        Me.grpStyles_mnu_Heading3Numbering_btn_off.Name = "grpStyles_mnu_Heading3Numbering_btn_off"
        Me.grpStyles_mnu_Heading3Numbering_btn_off.OfficeImageId = "BevelShapeGallery"
        Me.grpStyles_mnu_Heading3Numbering_btn_off.ShowImage = True
        '
        'grpStyles_NoNum
        '
        Me.grpStyles_NoNum.Items.Add(Me.grpStylesRpt_HeadingNoNum_StyleSet)
        Me.grpStyles_NoNum.Items.Add(Me.grpStylesRpt_Heading2NoNum_Rpt)
        Me.grpStyles_NoNum.Items.Add(Me.grpStylesRpt_Heading3NoNum_Rpt)
        Me.grpStyles_NoNum.Items.Add(Me.grpStylesRpt_Heading4NoNum_Rpt)
        Me.grpStyles_NoNum.Items.Add(Me.grpStylesRpt_Heading5NoNum_Rpt)
        Me.grpStyles_NoNum.Label = "Use Anywhere"
        Me.grpStyles_NoNum.Name = "grpStyles_NoNum"
        '
        'grpStylesRpt_HeadingNoNum_StyleSet
        '
        Me.grpStylesRpt_HeadingNoNum_StyleSet.KeyTip = "!N"
        Me.grpStylesRpt_HeadingNoNum_StyleSet.Label = "No # StyleSet"
        Me.grpStylesRpt_HeadingNoNum_StyleSet.Name = "grpStylesRpt_HeadingNoNum_StyleSet"
        Me.grpStylesRpt_HeadingNoNum_StyleSet.OfficeImageId = "FunctionsLogicalInsertGallery"
        Me.grpStylesRpt_HeadingNoNum_StyleSet.ShowImage = True
        Me.grpStylesRpt_HeadingNoNum_StyleSet.SuperTip = """This button will insert a report skeleton that uses most of the major 'no number" &
    "' report styles. These styles are not picked up by the TOC"""
        '
        'grpStylesRpt_Heading2NoNum_Rpt
        '
        Me.grpStylesRpt_Heading2NoNum_Rpt.KeyTip = "!2"
        Me.grpStylesRpt_Heading2NoNum_Rpt.Label = "H2 (no #)"
        Me.grpStylesRpt_Heading2NoNum_Rpt.Name = "grpStylesRpt_Heading2NoNum_Rpt"
        Me.grpStylesRpt_Heading2NoNum_Rpt.OfficeImageId = "BevelShapeGallery"
        Me.grpStylesRpt_Heading2NoNum_Rpt.ShowImage = True
        '
        'grpStylesRpt_Heading3NoNum_Rpt
        '
        Me.grpStylesRpt_Heading3NoNum_Rpt.KeyTip = "!3"
        Me.grpStylesRpt_Heading3NoNum_Rpt.Label = "H3 (no #)"
        Me.grpStylesRpt_Heading3NoNum_Rpt.Name = "grpStylesRpt_Heading3NoNum_Rpt"
        Me.grpStylesRpt_Heading3NoNum_Rpt.OfficeImageId = "BevelShapeGallery"
        Me.grpStylesRpt_Heading3NoNum_Rpt.ShowImage = True
        '
        'grpStylesRpt_Heading4NoNum_Rpt
        '
        Me.grpStylesRpt_Heading4NoNum_Rpt.KeyTip = "!4"
        Me.grpStylesRpt_Heading4NoNum_Rpt.Label = "H4 (no #)"
        Me.grpStylesRpt_Heading4NoNum_Rpt.Name = "grpStylesRpt_Heading4NoNum_Rpt"
        Me.grpStylesRpt_Heading4NoNum_Rpt.OfficeImageId = "BevelShapeGallery"
        Me.grpStylesRpt_Heading4NoNum_Rpt.ShowImage = True
        '
        'grpStylesRpt_Heading5NoNum_Rpt
        '
        Me.grpStylesRpt_Heading5NoNum_Rpt.KeyTip = "!5"
        Me.grpStylesRpt_Heading5NoNum_Rpt.Label = "H5 (no #)"
        Me.grpStylesRpt_Heading5NoNum_Rpt.Name = "grpStylesRpt_Heading5NoNum_Rpt"
        Me.grpStylesRpt_Heading5NoNum_Rpt.OfficeImageId = "BevelShapeGallery"
        Me.grpStylesRpt_Heading5NoNum_Rpt.ShowImage = True
        '
        'grpStyles_Appendices
        '
        Me.grpStyles_Appendices.Items.Add(Me.grpStylesApp_StyleSet)
        Me.grpStyles_Appendices.Items.Add(Me.grpStylesApp_Heading1_App)
        Me.grpStyles_Appendices.Items.Add(Me.grpStylesApp_Heading2_App)
        Me.grpStyles_Appendices.Items.Add(Me.grpStylesApp_Heading3_App)
        Me.grpStyles_Appendices.Items.Add(Me.grpStylesApp_Heading4_App)
        Me.grpStyles_Appendices.Items.Add(Me.grpStylesApp_Heading5_App)
        Me.grpStyles_Appendices.Label = "Appendix / Attachment"
        Me.grpStyles_Appendices.Name = "grpStyles_Appendices"
        '
        'grpStylesApp_StyleSet
        '
        Me.grpStylesApp_StyleSet.KeyTip = "AS"
        Me.grpStylesApp_StyleSet.Label = "App StyleSet"
        Me.grpStylesApp_StyleSet.Name = "grpStylesApp_StyleSet"
        Me.grpStylesApp_StyleSet.OfficeImageId = "FunctionsLogicalInsertGallery"
        Me.grpStylesApp_StyleSet.ScreenTip = "App StyleSet"
        Me.grpStylesApp_StyleSet.ShowImage = True
        Me.grpStylesApp_StyleSet.SuperTip = "This button will insert a skeleton that uses most of the major appendix/attachmen" &
    "t styles"
        '
        'grpStylesApp_Heading1_App
        '
        Me.grpStylesApp_Heading1_App.KeyTip = "A1"
        Me.grpStylesApp_Heading1_App.Label = "App H1"
        Me.grpStylesApp_Heading1_App.Name = "grpStylesApp_Heading1_App"
        Me.grpStylesApp_Heading1_App.OfficeImageId = "BevelShapeGallery"
        Me.grpStylesApp_Heading1_App.ScreenTip = "App H1"
        Me.grpStylesApp_Heading1_App.ShowImage = True
        Me.grpStylesApp_Heading1_App.SuperTip = "Appendix Heading 1 style - use in Appendix/Attachment only - number is prefixed b" &
    "y appendix/attachment #"
        Me.grpStylesApp_Heading1_App.Visible = False
        '
        'grpStylesApp_Heading2_App
        '
        Me.grpStylesApp_Heading2_App.KeyTip = "A2"
        Me.grpStylesApp_Heading2_App.Label = "App H2"
        Me.grpStylesApp_Heading2_App.Name = "grpStylesApp_Heading2_App"
        Me.grpStylesApp_Heading2_App.OfficeImageId = "BevelShapeGallery"
        Me.grpStylesApp_Heading2_App.ScreenTip = "App H2"
        Me.grpStylesApp_Heading2_App.ShowImage = True
        Me.grpStylesApp_Heading2_App.SuperTip = "Appendix Heading 2 style - use in Appendix/Attachment only - number is prefixed b" &
    "y appendix/attachment #"
        '
        'grpStylesApp_Heading3_App
        '
        Me.grpStylesApp_Heading3_App.KeyTip = "A3"
        Me.grpStylesApp_Heading3_App.Label = "App H3"
        Me.grpStylesApp_Heading3_App.Name = "grpStylesApp_Heading3_App"
        Me.grpStylesApp_Heading3_App.OfficeImageId = "BevelShapeGallery"
        Me.grpStylesApp_Heading3_App.ScreenTip = "App H3"
        Me.grpStylesApp_Heading3_App.ShowImage = True
        Me.grpStylesApp_Heading3_App.SuperTip = "Appendix Heading 3 style - use in Appendix/Attachment only - number is prefixed b" &
    "y appendix/attachment #"
        '
        'grpStylesApp_Heading4_App
        '
        Me.grpStylesApp_Heading4_App.KeyTip = "A4"
        Me.grpStylesApp_Heading4_App.Label = "App H4"
        Me.grpStylesApp_Heading4_App.Name = "grpStylesApp_Heading4_App"
        Me.grpStylesApp_Heading4_App.OfficeImageId = "BevelShapeGallery"
        Me.grpStylesApp_Heading4_App.ScreenTip = "App H4"
        Me.grpStylesApp_Heading4_App.ShowImage = True
        Me.grpStylesApp_Heading4_App.SuperTip = "Appendix Heading 4 style - use in Appendix/Attachment only"
        '
        'grpStylesApp_Heading5_App
        '
        Me.grpStylesApp_Heading5_App.KeyTip = "A5"
        Me.grpStylesApp_Heading5_App.Label = "App H5"
        Me.grpStylesApp_Heading5_App.Name = "grpStylesApp_Heading5_App"
        Me.grpStylesApp_Heading5_App.OfficeImageId = "BevelShapeGallery"
        Me.grpStylesApp_Heading5_App.ScreenTip = "App H5"
        Me.grpStylesApp_Heading5_App.ShowImage = True
        Me.grpStylesApp_Heading5_App.SuperTip = "Appendix Heading 5 style - use in Appendix/Attachment only"
        Me.grpStylesApp_Heading5_App.Tag = ""
        '
        'grpStyles_Text
        '
        Me.grpStyles_Text.Items.Add(Me.grpStylesText_BodyText)
        Me.grpStyles_Text.Items.Add(Me.grpStylesRpt_Intro)
        Me.grpStyles_Text.Items.Add(Me.grpStylesOther_Quote)
        Me.grpStyles_Text.Items.Add(Me.grpStylesOther_QuoteBlt)
        Me.grpStyles_Text.Items.Add(Me.grpStylesOther_QuoteSource)
        Me.grpStyles_Text.Label = "Text"
        Me.grpStyles_Text.Name = "grpStyles_Text"
        '
        'grpStylesText_BodyText
        '
        Me.grpStylesText_BodyText.KeyTip = "TB"
        Me.grpStylesText_BodyText.Label = "Body Text"
        Me.grpStylesText_BodyText.Name = "grpStylesText_BodyText"
        Me.grpStylesText_BodyText.OfficeImageId = "ContentControlText"
        Me.grpStylesText_BodyText.ScreenTip = "Body Text"
        Me.grpStylesText_BodyText.ShowImage = True
        Me.grpStylesText_BodyText.SuperTip = "Body Text style is the style for all basic text."
        '
        'grpStylesRpt_Intro
        '
        Me.grpStylesRpt_Intro.KeyTip = "TI"
        Me.grpStylesRpt_Intro.Label = "Intro/Summary"
        Me.grpStylesRpt_Intro.Name = "grpStylesRpt_Intro"
        Me.grpStylesRpt_Intro.OfficeImageId = "StartOfDocument"
        Me.grpStylesRpt_Intro.ScreenTip = "Intro/Summary"
        Me.grpStylesRpt_Intro.ShowImage = True
        Me.grpStylesRpt_Intro.SuperTip = "Use this style for the introduction/summary following an Executive Summary headin" &
    "g or a Chapter heading"
        '
        'grpStylesOther_Quote
        '
        Me.grpStylesOther_Quote.KeyTip = "TQQ"
        Me.grpStylesOther_Quote.Label = "Quote"
        Me.grpStylesOther_Quote.Name = "grpStylesOther_Quote"
        Me.grpStylesOther_Quote.OfficeImageId = "SparklineWinLossInsert"
        Me.grpStylesOther_Quote.ScreenTip = "Quote"
        Me.grpStylesOther_Quote.ShowImage = True
        Me.grpStylesOther_Quote.SuperTip = "Quote style - applies smaller font size"
        '
        'grpStylesOther_QuoteBlt
        '
        Me.grpStylesOther_QuoteBlt.KeyTip = "TQB"
        Me.grpStylesOther_QuoteBlt.Label = "Qte Blt"
        Me.grpStylesOther_QuoteBlt.Name = "grpStylesOther_QuoteBlt"
        Me.grpStylesOther_QuoteBlt.OfficeImageId = "SparklineCustomWeight"
        Me.grpStylesOther_QuoteBlt.ScreenTip = "Qte Blt"
        Me.grpStylesOther_QuoteBlt.ShowImage = True
        Me.grpStylesOther_QuoteBlt.SuperTip = "Quote List Bullet style - apply to bulleted points within a quote."
        '
        'grpStylesOther_QuoteSource
        '
        Me.grpStylesOther_QuoteSource.KeyTip = "TQS"
        Me.grpStylesOther_QuoteSource.Label = "Qte Source"
        Me.grpStylesOther_QuoteSource.Name = "grpStylesOther_QuoteSource"
        Me.grpStylesOther_QuoteSource.OfficeImageId = "SparklineCustomWeight"
        Me.grpStylesOther_QuoteSource.ScreenTip = "Qte Source"
        Me.grpStylesOther_QuoteSource.ShowImage = True
        Me.grpStylesOther_QuoteSource.SuperTip = "Quote source style - right aligns and italicizes quote source text"
        '
        'grpStyles_Lists
        '
        Me.grpStyles_Lists.Items.Add(Me.grpStylesLists_List1)
        Me.grpStyles_Lists.Items.Add(Me.grpStylesLists_List2)
        Me.grpStyles_Lists.Items.Add(Me.grpStylesLists_List3)
        Me.grpStyles_Lists.Items.Add(Me.Separator55)
        Me.grpStyles_Lists.Items.Add(Me.grpStylesLists_ListNumber1)
        Me.grpStyles_Lists.Items.Add(Me.grpStylesLists_ListNumber2)
        Me.grpStyles_Lists.Items.Add(Me.grpStylesLists_ListNumber3)
        Me.grpStyles_Lists.Label = "Lists"
        Me.grpStyles_Lists.Name = "grpStyles_Lists"
        '
        'grpStylesLists_List1
        '
        Me.grpStylesLists_List1.KeyTip = "L1"
        Me.grpStylesLists_List1.Label = "List Bullet 1"
        Me.grpStylesLists_List1.Name = "grpStylesLists_List1"
        Me.grpStylesLists_List1.OfficeImageId = "Bullets"
        Me.grpStylesLists_List1.ScreenTip = "List Bullet 1"
        Me.grpStylesLists_List1.ShowImage = True
        '
        'grpStylesLists_List2
        '
        Me.grpStylesLists_List2.KeyTip = "L2"
        Me.grpStylesLists_List2.Label = "List Bullet 2"
        Me.grpStylesLists_List2.Name = "grpStylesLists_List2"
        Me.grpStylesLists_List2.OfficeImageId = "Bullets"
        Me.grpStylesLists_List2.ScreenTip = "List Bullet 2"
        Me.grpStylesLists_List2.ShowImage = True
        Me.grpStylesLists_List2.SuperTip = "List bullet level 2 style for the body of the report"
        '
        'grpStylesLists_List3
        '
        Me.grpStylesLists_List3.KeyTip = "L3"
        Me.grpStylesLists_List3.Label = "List Bullet 3"
        Me.grpStylesLists_List3.Name = "grpStylesLists_List3"
        Me.grpStylesLists_List3.OfficeImageId = "Bullets"
        Me.grpStylesLists_List3.ScreenTip = "List Bullet 3"
        Me.grpStylesLists_List3.ShowImage = True
        Me.grpStylesLists_List3.SuperTip = "List bullet level 3 style for the body of the report"
        '
        'Separator55
        '
        Me.Separator55.Name = "Separator55"
        '
        'grpStylesLists_ListNumber1
        '
        Me.grpStylesLists_ListNumber1.KeyTip = "N1"
        Me.grpStylesLists_ListNumber1.Label = "List Number 1"
        Me.grpStylesLists_ListNumber1.Name = "grpStylesLists_ListNumber1"
        Me.grpStylesLists_ListNumber1.OfficeImageId = "Numbering"
        Me.grpStylesLists_ListNumber1.ScreenTip = "List Number 1"
        Me.grpStylesLists_ListNumber1.ShowImage = True
        Me.grpStylesLists_ListNumber1.SuperTip = "List number style"
        '
        'grpStylesLists_ListNumber2
        '
        Me.grpStylesLists_ListNumber2.KeyTip = "N2"
        Me.grpStylesLists_ListNumber2.Label = "List Number 2"
        Me.grpStylesLists_ListNumber2.Name = "grpStylesLists_ListNumber2"
        Me.grpStylesLists_ListNumber2.OfficeImageId = "Numbering"
        Me.grpStylesLists_ListNumber2.ScreenTip = "List Number 2"
        Me.grpStylesLists_ListNumber2.ShowImage = True
        Me.grpStylesLists_ListNumber2.SuperTip = "List number 2 style"
        '
        'grpStylesLists_ListNumber3
        '
        Me.grpStylesLists_ListNumber3.KeyTip = "N3"
        Me.grpStylesLists_ListNumber3.Label = "List Number 3"
        Me.grpStylesLists_ListNumber3.Name = "grpStylesLists_ListNumber3"
        Me.grpStylesLists_ListNumber3.OfficeImageId = "Numbering"
        Me.grpStylesLists_ListNumber3.ScreenTip = "List Number 3"
        Me.grpStylesLists_ListNumber3.ShowImage = True
        Me.grpStylesLists_ListNumber3.SuperTip = "List number 3 style"
        '
        'grpStyles_Emphasis
        '
        Me.grpStyles_Emphasis.Items.Add(Me.tbStyles_mnu_Emphasis)
        Me.grpStyles_Emphasis.Label = "Emphasis"
        Me.grpStyles_Emphasis.Name = "grpStyles_Emphasis"
        '
        'tbStyles_mnu_Emphasis
        '
        Me.tbStyles_mnu_Emphasis.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.tbStyles_mnu_Emphasis.Items.Add(Me.grpPullouts_emphasisBox_TextStyle_Left_2)
        Me.tbStyles_mnu_Emphasis.Items.Add(Me.grpPullouts_emphasisBox_TextStyle_Centre_2)
        Me.tbStyles_mnu_Emphasis.Items.Add(Me.grpPullouts_emphasisBox_TextStyle_Right_2)
        Me.tbStyles_mnu_Emphasis.KeyTip = "MP"
        Me.tbStyles_mnu_Emphasis.Label = "Emphasis"
        Me.tbStyles_mnu_Emphasis.Name = "tbStyles_mnu_Emphasis"
        Me.tbStyles_mnu_Emphasis.OfficeImageId = "ControlLogo"
        Me.tbStyles_mnu_Emphasis.ScreenTip = "Emphasis"
        Me.tbStyles_mnu_Emphasis.ShowImage = True
        Me.tbStyles_mnu_Emphasis.SuperTip = "Select various Emphasis boxes to highlight important issues.. Do not use the 'box" &
    "' options in documents that need to be 'Accessible'. For those documents you wil" &
    "l need to use the text styles"
        '
        'grpPullouts_emphasisBox_TextStyle_Left_2
        '
        Me.grpPullouts_emphasisBox_TextStyle_Left_2.Label = "Emphasis Text (&Left) Style"
        Me.grpPullouts_emphasisBox_TextStyle_Left_2.Name = "grpPullouts_emphasisBox_TextStyle_Left_2"
        Me.grpPullouts_emphasisBox_TextStyle_Left_2.OfficeImageId = "ControlLogo"
        Me.grpPullouts_emphasisBox_TextStyle_Left_2.ScreenTip = "Emphasis Text (Left)"
        Me.grpPullouts_emphasisBox_TextStyle_Left_2.ShowImage = True
        Me.grpPullouts_emphasisBox_TextStyle_Left_2.SuperTip = "Will apply the 'Emphasis Text (Left)' style to the selected paragraph. This style" &
    " is left justified... This can be used in documents that need to be 'Accessible'" &
    ""
        '
        'grpPullouts_emphasisBox_TextStyle_Centre_2
        '
        Me.grpPullouts_emphasisBox_TextStyle_Centre_2.Label = "Emphasis Text (&Centre) Style"
        Me.grpPullouts_emphasisBox_TextStyle_Centre_2.Name = "grpPullouts_emphasisBox_TextStyle_Centre_2"
        Me.grpPullouts_emphasisBox_TextStyle_Centre_2.OfficeImageId = "ControlLogo"
        Me.grpPullouts_emphasisBox_TextStyle_Centre_2.ScreenTip = "Emphasis Text (Centre)"
        Me.grpPullouts_emphasisBox_TextStyle_Centre_2.ShowImage = True
        Me.grpPullouts_emphasisBox_TextStyle_Centre_2.SuperTip = "Will apply the 'Emphasis Text (Centre)' style to the selected paragraph. This sty" &
    "le is centre justified... This can be used in documents that need to be 'Accessi" &
    "ble'"
        '
        'grpPullouts_emphasisBox_TextStyle_Right_2
        '
        Me.grpPullouts_emphasisBox_TextStyle_Right_2.Label = "Emphasis Text (&Right) Style"
        Me.grpPullouts_emphasisBox_TextStyle_Right_2.Name = "grpPullouts_emphasisBox_TextStyle_Right_2"
        Me.grpPullouts_emphasisBox_TextStyle_Right_2.OfficeImageId = "ControlLogo"
        Me.grpPullouts_emphasisBox_TextStyle_Right_2.ScreenTip = "Emphasis Text (Right) "
        Me.grpPullouts_emphasisBox_TextStyle_Right_2.ShowImage = True
        Me.grpPullouts_emphasisBox_TextStyle_Right_2.SuperTip = "Will apply the 'Emphasis Text (Right)' style to the selected paragraph. This styl" &
    "e is right justified... This can be used in documents that need to be 'Accessibl" &
    "e'"
        '
        'grpStyles_resetStyles
        '
        Me.grpStyles_resetStyles.Items.Add(Me.grpStylesTools_to_PrintDefault)
        Me.grpStyles_resetStyles.Items.Add(Me.grpStylesTools_to_DisplayDefault)
        Me.grpStyles_resetStyles.Items.Add(Me.tbStyles_grpResetStyles_mnu_ResetStyles)
        Me.grpStyles_resetStyles.Label = "Reset Styles"
        Me.grpStyles_resetStyles.Name = "grpStyles_resetStyles"
        '
        'grpStylesTools_to_PrintDefault
        '
        Me.grpStylesTools_to_PrintDefault.Image = Global.AA_GeneralReport_Addin.My.Resources.Resources.icons_colour_Purple_Print
        Me.grpStylesTools_to_PrintDefault.KeyTip = "XP"
        Me.grpStylesTools_to_PrintDefault.Label = "Print Colour"
        Me.grpStylesTools_to_PrintDefault.Name = "grpStylesTools_to_PrintDefault"
        Me.grpStylesTools_to_PrintDefault.ScreenTip = "Print Colour"
        Me.grpStylesTools_to_PrintDefault.ShowImage = True
        Me.grpStylesTools_to_PrintDefault.SuperTip = "This button will set the font colour in the Cover Page and Back Contacts Page to " &
    "the light purple Print colour. If you find this difficult on the eyes, then sele" &
    "ct the Display Colour"
        '
        'grpStylesTools_to_DisplayDefault
        '
        Me.grpStylesTools_to_DisplayDefault.Image = Global.AA_GeneralReport_Addin.My.Resources.Resources.icons_colour_Purple_Display
        Me.grpStylesTools_to_DisplayDefault.KeyTip = "XD"
        Me.grpStylesTools_to_DisplayDefault.Label = "Display Colour"
        Me.grpStylesTools_to_DisplayDefault.Name = "grpStylesTools_to_DisplayDefault"
        Me.grpStylesTools_to_DisplayDefault.ScreenTip = "Display Colour"
        Me.grpStylesTools_to_DisplayDefault.ShowImage = True
        Me.grpStylesTools_to_DisplayDefault.SuperTip = resources.GetString("grpStylesTools_to_DisplayDefault.SuperTip")
        '
        'tbStyles_grpResetStyles_mnu_ResetStyles
        '
        Me.tbStyles_grpResetStyles_mnu_ResetStyles.Items.Add(Me.tabStyles_btn_resetStylesForRptPrt)
        Me.tbStyles_grpResetStyles_mnu_ResetStyles.Items.Add(Me.tabStyles_btn_resetStylesForRptLnd)
        Me.tbStyles_grpResetStyles_mnu_ResetStyles.Items.Add(Me.tabStyles_btn_resetStylesForRptBrf)
        Me.tbStyles_grpResetStyles_mnu_ResetStyles.KeyTip = "XR"
        Me.tbStyles_grpResetStyles_mnu_ResetStyles.Label = "Reset Report Styles"
        Me.tbStyles_grpResetStyles_mnu_ResetStyles.Name = "tbStyles_grpResetStyles_mnu_ResetStyles"
        Me.tbStyles_grpResetStyles_mnu_ResetStyles.OfficeImageId = "BevelShapeGallery"
        Me.tbStyles_grpResetStyles_mnu_ResetStyles.ScreenTip = "Reset Styles"
        Me.tbStyles_grpResetStyles_mnu_ResetStyles.ShowImage = True
        Me.tbStyles_grpResetStyles_mnu_ResetStyles.SuperTip = "Provides function to reset the styles in the current document to those required b" &
    "y the various ACIL Allen report types"
        '
        'tabStyles_btn_resetStylesForRptPrt
        '
        Me.tabStyles_btn_resetStylesForRptPrt.Label = "Reset for &Portrait Report"
        Me.tabStyles_btn_resetStylesForRptPrt.Name = "tabStyles_btn_resetStylesForRptPrt"
        Me.tabStyles_btn_resetStylesForRptPrt.OfficeImageId = "BevelShapeGallery"
        Me.tabStyles_btn_resetStylesForRptPrt.ShowImage = True
        '
        'tabStyles_btn_resetStylesForRptLnd
        '
        Me.tabStyles_btn_resetStylesForRptLnd.Label = "Reset for &Landscape Report"
        Me.tabStyles_btn_resetStylesForRptLnd.Name = "tabStyles_btn_resetStylesForRptLnd"
        Me.tabStyles_btn_resetStylesForRptLnd.OfficeImageId = "BevelShapeGallery"
        Me.tabStyles_btn_resetStylesForRptLnd.ShowImage = True
        '
        'tabStyles_btn_resetStylesForRptBrf
        '
        Me.tabStyles_btn_resetStylesForRptBrf.Label = "Reset for &Brief Report"
        Me.tabStyles_btn_resetStylesForRptBrf.Name = "tabStyles_btn_resetStylesForRptBrf"
        Me.tabStyles_btn_resetStylesForRptBrf.OfficeImageId = "BevelShapeGallery"
        Me.tabStyles_btn_resetStylesForRptBrf.ShowImage = True
        '
        'grpStyles_resetCaptions
        '
        Me.grpStyles_resetCaptions.Items.Add(Me.grpStylesTools_resetCaptions)
        Me.grpStyles_resetCaptions.Label = "Reset Captions"
        Me.grpStyles_resetCaptions.Name = "grpStyles_resetCaptions"
        '
        'grpStylesTools_resetCaptions
        '
        Me.grpStylesTools_resetCaptions.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grpStylesTools_resetCaptions.KeyTip = "XT"
        Me.grpStylesTools_resetCaptions.Label = "Reset captions"
        Me.grpStylesTools_resetCaptions.Name = "grpStylesTools_resetCaptions"
        Me.grpStylesTools_resetCaptions.OfficeImageId = "BevelShapeGallery"
        Me.grpStylesTools_resetCaptions.ScreenTip = "Reset captions"
        Me.grpStylesTools_resetCaptions.ShowImage = True
        Me.grpStylesTools_resetCaptions.SuperTip = "This function refreshes the Acil Allen custom captions in the current document. U" &
    "se this function if you think you are missing a caption that you want to use for" &
    " cross referencing"
        '
        'tab_aa_Placeholders
        '
        Me.tab_aa_Placeholders.Groups.Add(Me.grp_PlaceHolders)
        Me.tab_aa_Placeholders.Groups.Add(Me.grp_special_AATableFormatting)
        Me.tab_aa_Placeholders.Groups.Add(Me.grp_floatingPlaceholders)
        Me.tab_aa_Placeholders.Groups.Add(Me.grp_Plh_miscPlaceholders)
        Me.tab_aa_Placeholders.KeyTip = "JL"
        Me.tab_aa_Placeholders.Label = "Placeholders"
        Me.tab_aa_Placeholders.Name = "tab_aa_Placeholders"
        Me.tab_aa_Placeholders.Position = Me.Factory.RibbonPosition.BeforeOfficeId("TabInsert")
        '
        'grp_PlaceHolders
        '
        Me.grp_PlaceHolders.Items.Add(Me.mnuCloseDocuments161)
        Me.grp_PlaceHolders.Items.Add(Me.mnuCloseDocuments2233)
        Me.grp_PlaceHolders.Items.Add(Me.mnuCloseDocuments1)
        Me.grp_PlaceHolders.Items.Add(Me.mnu_grpBoxes_Recommendations)
        Me.grp_PlaceHolders.Items.Add(Me.mnu_grpBoxes_Findings)
        Me.grp_PlaceHolders.Items.Add(Me.grpPullouts_mnu01)
        Me.grp_PlaceHolders.Items.Add(Me.grpReport_mnu_CaseStudies)
        Me.grp_PlaceHolders.Items.Add(Me.Separator37)
        Me.grp_PlaceHolders.Items.Add(Me.mnuCloseDocuments16)
        Me.grp_PlaceHolders.Items.Add(Me.mnuCloseDocuments33)
        Me.grp_PlaceHolders.Items.Add(Me.grpStylesRpt_mnu_tbls_00)
        Me.grp_PlaceHolders.Items.Add(Me.Separator38)
        Me.grp_PlaceHolders.Items.Add(Me.grpPlh_btn_buildCustomTable)
        Me.grp_PlaceHolders.Items.Add(Me.grpTbl_Styles)
        Me.grp_PlaceHolders.Items.Add(Me.grpTbls_TableTextStyle)
        Me.grp_PlaceHolders.Items.Add(Me.mnuCloseDocuments4)
        Me.grp_PlaceHolders.Items.Add(Me.grpTbls_AllStyles_small)
        Me.grp_PlaceHolders.Items.Add(Me.grpTbls_TableTextStyle_small)
        Me.grp_PlaceHolders.Items.Add(Me.grpPlh_mnu_TblPlaceholders)
        Me.grp_PlaceHolders.Items.Add(Me.grpPlh_mnu_SourceAndNote)
        Me.grp_PlaceHolders.Items.Add(Me.grpPlh_mnu_DeleteTable)
        Me.grp_PlaceHolders.Label = "Placeholders"
        Me.grp_PlaceHolders.Name = "grp_PlaceHolders"
        '
        'mnuCloseDocuments161
        '
        Me.mnuCloseDocuments161.Items.Add(Me.grpBoxes_Box)
        Me.mnuCloseDocuments161.Items.Add(Me.grpBoxes_AppendixBox)
        Me.mnuCloseDocuments161.Items.Add(Me.grpBoxes_ESBox)
        Me.mnuCloseDocuments161.Items.Add(Me.grpBoxes_LTBox)
        Me.mnuCloseDocuments161.Items.Add(Me.Separator31)
        Me.mnuCloseDocuments161.Items.Add(Me.grpBoxes_CaptionAndHeading)
        Me.mnuCloseDocuments161.Items.Add(Me.grpBoxes_CaptionAndHeadingES)
        Me.mnuCloseDocuments161.Items.Add(Me.grpBoxes_CaptionAndHeadingApp)
        Me.mnuCloseDocuments161.KeyTip = "BB"
        Me.mnuCloseDocuments161.Label = "Boxes"
        Me.mnuCloseDocuments161.Name = "mnuCloseDocuments161"
        Me.mnuCloseDocuments161.OfficeImageId = "ControlLogo"
        Me.mnuCloseDocuments161.ScreenTip = "Boxes"
        Me.mnuCloseDocuments161.ShowImage = True
        Me.mnuCloseDocuments161.SuperTip = """Select from the list of preformatted sized placeholders which have the caption l" &
    "abel Box."""
        '
        'grpBoxes_Box
        '
        Me.grpBoxes_Box.Label = "Box (&Report)"
        Me.grpBoxes_Box.Name = "grpBoxes_Box"
        Me.grpBoxes_Box.OfficeImageId = "ControlLogo"
        Me.grpBoxes_Box.ScreenTip = "Box Report"
        Me.grpBoxes_Box.ShowImage = True
        Me.grpBoxes_Box.SuperTip = """Inserts a box to fit within the standard text margin. """
        '
        'grpBoxes_AppendixBox
        '
        Me.grpBoxes_AppendixBox.Label = "Box (&Appendix)"
        Me.grpBoxes_AppendixBox.Name = "grpBoxes_AppendixBox"
        Me.grpBoxes_AppendixBox.OfficeImageId = "ControlLogo"
        Me.grpBoxes_AppendixBox.ScreenTip = "Box Appendix"
        Me.grpBoxes_AppendixBox.ShowImage = True
        Me.grpBoxes_AppendixBox.SuperTip = """Inserts a box to fit within the standard text margin.  Box number is prefixed by" &
    " Appendix number. """
        '
        'grpBoxes_ESBox
        '
        Me.grpBoxes_ESBox.Label = "Box (&ES)"
        Me.grpBoxes_ESBox.Name = "grpBoxes_ESBox"
        Me.grpBoxes_ESBox.OfficeImageId = "ControlLogo"
        Me.grpBoxes_ESBox.ScreenTip = "Box ES"
        Me.grpBoxes_ESBox.ShowImage = True
        Me.grpBoxes_ESBox.SuperTip = """Inserts a box to fit within the standard text margin with number prefixed by ES." &
    " """
        '
        'grpBoxes_LTBox
        '
        Me.grpBoxes_LTBox.Label = "Box (&LT for letters, memos etc)"
        Me.grpBoxes_LTBox.Name = "grpBoxes_LTBox"
        Me.grpBoxes_LTBox.OfficeImageId = "ControlLogo"
        Me.grpBoxes_LTBox.ScreenTip = "Box LT"
        Me.grpBoxes_LTBox.ShowImage = True
        Me.grpBoxes_LTBox.SuperTip = """To be used for letters, memos etc. This menu item inserts a Box placeholder that" &
    " uses 'Stationery Numbering' for the Caption."""
        '
        'Separator31
        '
        Me.Separator31.Name = "Separator31"
        '
        'grpBoxes_CaptionAndHeading
        '
        Me.grpBoxes_CaptionAndHeading.Label = "Insert Box Caption and Heading only"
        Me.grpBoxes_CaptionAndHeading.Name = "grpBoxes_CaptionAndHeading"
        Me.grpBoxes_CaptionAndHeading.OfficeImageId = "AdvertisePublishAs"
        Me.grpBoxes_CaptionAndHeading.ScreenTip = "Base Caption and Heading"
        Me.grpBoxes_CaptionAndHeading.ShowImage = True
        Me.grpBoxes_CaptionAndHeading.SuperTip = """Inserts Box caption and number with heading to be overtyped.  Place cursor in fi" &
    "rst cell of table, before running this macro."""
        '
        'grpBoxes_CaptionAndHeadingES
        '
        Me.grpBoxes_CaptionAndHeadingES.Label = "Insert ES Box Caption and Heading only"
        Me.grpBoxes_CaptionAndHeadingES.Name = "grpBoxes_CaptionAndHeadingES"
        Me.grpBoxes_CaptionAndHeadingES.OfficeImageId = "AdvertisePublishAs"
        Me.grpBoxes_CaptionAndHeadingES.ScreenTip = "ES Caption and Heading"
        Me.grpBoxes_CaptionAndHeadingES.ShowImage = True
        Me.grpBoxes_CaptionAndHeadingES.SuperTip = """Inserts ES prefixed Box caption and number for use in Executive Summary.  Place " &
    "cursor in first cell of table, before running this macro."""
        '
        'grpBoxes_CaptionAndHeadingApp
        '
        Me.grpBoxes_CaptionAndHeadingApp.Label = "Insert AP Box Caption and Heading only"
        Me.grpBoxes_CaptionAndHeadingApp.Name = "grpBoxes_CaptionAndHeadingApp"
        Me.grpBoxes_CaptionAndHeadingApp.OfficeImageId = "AdvertisePublishAs"
        Me.grpBoxes_CaptionAndHeadingApp.ScreenTip = "Appendix Caption and Heading"
        Me.grpBoxes_CaptionAndHeadingApp.ShowImage = True
        Me.grpBoxes_CaptionAndHeadingApp.SuperTip = """Inserts AP prefixed Box caption and number. Use in an Appendix. Place cursor in " &
    "first cell of table, before running this macro."""
        '
        'mnuCloseDocuments2233
        '
        Me.mnuCloseDocuments2233.Items.Add(Me.grpBoxes_BoxTextBoldItalic)
        Me.mnuCloseDocuments2233.Items.Add(Me.grpBoxes_BoxText)
        Me.mnuCloseDocuments2233.Items.Add(Me.grpBoxes_SideHeading1)
        Me.mnuCloseDocuments2233.Items.Add(Me.grpBoxes_SideHeading2)
        Me.mnuCloseDocuments2233.Items.Add(Me.Separator32)
        Me.mnuCloseDocuments2233.Items.Add(Me.grpBoxes_BoxListBullet)
        Me.mnuCloseDocuments2233.Items.Add(Me.grpBoxes_BoxListBullet2)
        Me.mnuCloseDocuments2233.Items.Add(Me.grpBoxes_BoxListBullet3)
        Me.mnuCloseDocuments2233.Items.Add(Me.Separator33)
        Me.mnuCloseDocuments2233.Items.Add(Me.grpBoxes_BoxListNumber)
        Me.mnuCloseDocuments2233.Items.Add(Me.grpBoxes_BoxListNumber2)
        Me.mnuCloseDocuments2233.Items.Add(Me.grpBoxes_BoxListNumber3)
        Me.mnuCloseDocuments2233.Items.Add(Me.Separator34)
        Me.mnuCloseDocuments2233.Items.Add(Me.grpBoxes_BoxQuote)
        Me.mnuCloseDocuments2233.Items.Add(Me.grpBoxes_BoxQuoteListBullet)
        Me.mnuCloseDocuments2233.Items.Add(Me.grpBoxes_BoxQuoteSource)
        Me.mnuCloseDocuments2233.Items.Add(Me.Separator35)
        Me.mnuCloseDocuments2233.Items.Add(Me.grpBoxes_boxContent_mnu)
        Me.mnuCloseDocuments2233.KeyTip = "BS"
        Me.mnuCloseDocuments2233.Label = "Box Styles"
        Me.mnuCloseDocuments2233.Name = "mnuCloseDocuments2233"
        Me.mnuCloseDocuments2233.OfficeImageId = "BevelShapeGallery"
        Me.mnuCloseDocuments2233.ScreenTip = "Box Styles"
        Me.mnuCloseDocuments2233.ShowImage = True
        Me.mnuCloseDocuments2233.SuperTip = """This menu provides text and bulleted/numbered list styles that can be applied to" &
    " text within a box. Also a macro to restart or continue a numbered list within a" &
    " box."""
        '
        'grpBoxes_BoxTextBoldItalic
        '
        Me.grpBoxes_BoxTextBoldItalic.Label = "Box text (Bold &Italic)"
        Me.grpBoxes_BoxTextBoldItalic.Name = "grpBoxes_BoxTextBoldItalic"
        Me.grpBoxes_BoxTextBoldItalic.OfficeImageId = "T"
        Me.grpBoxes_BoxTextBoldItalic.ScreenTip = "Box text (Bold Italic)"
        Me.grpBoxes_BoxTextBoldItalic.ShowImage = True
        Me.grpBoxes_BoxTextBoldItalic.SuperTip = """Applies Box Text (Bold Italic) style to selection"""
        '
        'grpBoxes_BoxText
        '
        Me.grpBoxes_BoxText.Label = "Box &text"
        Me.grpBoxes_BoxText.Name = "grpBoxes_BoxText"
        Me.grpBoxes_BoxText.OfficeImageId = "T"
        Me.grpBoxes_BoxText.ShowImage = True
        '
        'grpBoxes_SideHeading1
        '
        Me.grpBoxes_SideHeading1.Label = "Box Side Heading &1"
        Me.grpBoxes_SideHeading1.Name = "grpBoxes_SideHeading1"
        Me.grpBoxes_SideHeading1.OfficeImageId = "T"
        Me.grpBoxes_SideHeading1.ScreenTip = "Box Side Heading 1"
        Me.grpBoxes_SideHeading1.ShowImage = True
        Me.grpBoxes_SideHeading1.SuperTip = """Applies Box Side Heading 1 style to selection. If applying this style to a headi" &
    "ng in a framed box, the box will split into separate boxes.  Drag through to sel" &
    "ect the whole box, then reframe."""
        '
        'grpBoxes_SideHeading2
        '
        Me.grpBoxes_SideHeading2.Label = "Box Side Heading &2"
        Me.grpBoxes_SideHeading2.Name = "grpBoxes_SideHeading2"
        Me.grpBoxes_SideHeading2.OfficeImageId = "T"
        Me.grpBoxes_SideHeading2.ScreenTip = "Box Side Heading 2"
        Me.grpBoxes_SideHeading2.ShowImage = True
        Me.grpBoxes_SideHeading2.SuperTip = """Applies Box Side Heading 2 style to selection. If applying this style to a headi" &
    "ng in a framed box, the box will split into separate boxes.  Drag through to sel" &
    "ect the whole box, then reframe."""
        '
        'Separator32
        '
        Me.Separator32.Name = "Separator32"
        '
        'grpBoxes_BoxListBullet
        '
        Me.grpBoxes_BoxListBullet.Label = "Box List Bullet"
        Me.grpBoxes_BoxListBullet.Name = "grpBoxes_BoxListBullet"
        Me.grpBoxes_BoxListBullet.OfficeImageId = "Bullets"
        Me.grpBoxes_BoxListBullet.ScreenTip = "Box List Bullet"
        Me.grpBoxes_BoxListBullet.ShowImage = True
        Me.grpBoxes_BoxListBullet.SuperTip = """Applies Box List Bullet style to selection"""
        '
        'grpBoxes_BoxListBullet2
        '
        Me.grpBoxes_BoxListBullet2.Label = "Box List Bullet 2"
        Me.grpBoxes_BoxListBullet2.Name = "grpBoxes_BoxListBullet2"
        Me.grpBoxes_BoxListBullet2.OfficeImageId = "Bullets"
        Me.grpBoxes_BoxListBullet2.ScreenTip = "Box List Bullet 2"
        Me.grpBoxes_BoxListBullet2.ShowImage = True
        Me.grpBoxes_BoxListBullet2.SuperTip = """Applies Box List Bullet 2 style to selection"""
        '
        'grpBoxes_BoxListBullet3
        '
        Me.grpBoxes_BoxListBullet3.Label = "Box List Bullet 3"
        Me.grpBoxes_BoxListBullet3.Name = "grpBoxes_BoxListBullet3"
        Me.grpBoxes_BoxListBullet3.OfficeImageId = "Bullets"
        Me.grpBoxes_BoxListBullet3.ScreenTip = "Box List Bullet 3"
        Me.grpBoxes_BoxListBullet3.ShowImage = True
        Me.grpBoxes_BoxListBullet3.SuperTip = """Applies Box List Bullet 3 style to selection"""
        '
        'Separator33
        '
        Me.Separator33.Name = "Separator33"
        '
        'grpBoxes_BoxListNumber
        '
        Me.grpBoxes_BoxListNumber.Label = "Box List Number style"
        Me.grpBoxes_BoxListNumber.Name = "grpBoxes_BoxListNumber"
        Me.grpBoxes_BoxListNumber.OfficeImageId = "Numbering"
        Me.grpBoxes_BoxListNumber.ScreenTip = "Box List Number style"
        Me.grpBoxes_BoxListNumber.ShowImage = True
        Me.grpBoxes_BoxListNumber.SuperTip = """Applies Box List Number style to selection"""
        '
        'grpBoxes_BoxListNumber2
        '
        Me.grpBoxes_BoxListNumber2.Label = "Box List Number 2 style"
        Me.grpBoxes_BoxListNumber2.Name = "grpBoxes_BoxListNumber2"
        Me.grpBoxes_BoxListNumber2.OfficeImageId = "Numbering"
        Me.grpBoxes_BoxListNumber2.ScreenTip = "Box List Number 2 style"
        Me.grpBoxes_BoxListNumber2.ShowImage = True
        Me.grpBoxes_BoxListNumber2.SuperTip = """Applies Box List Number 2 style to selection"""
        '
        'grpBoxes_BoxListNumber3
        '
        Me.grpBoxes_BoxListNumber3.Label = "Box List Number 3 style"
        Me.grpBoxes_BoxListNumber3.Name = "grpBoxes_BoxListNumber3"
        Me.grpBoxes_BoxListNumber3.OfficeImageId = "Numbering"
        Me.grpBoxes_BoxListNumber3.ScreenTip = "Box List Number 3 style"
        Me.grpBoxes_BoxListNumber3.ShowImage = True
        Me.grpBoxes_BoxListNumber3.SuperTip = """Applies Box List Number 3 style to selection"""
        '
        'Separator34
        '
        Me.Separator34.Name = "Separator34"
        '
        'grpBoxes_BoxQuote
        '
        Me.grpBoxes_BoxQuote.Label = "Box Quote style"
        Me.grpBoxes_BoxQuote.Name = "grpBoxes_BoxQuote"
        Me.grpBoxes_BoxQuote.OfficeImageId = "Q"
        Me.grpBoxes_BoxQuote.ScreenTip = "Box Quote style"
        Me.grpBoxes_BoxQuote.ShowImage = True
        Me.grpBoxes_BoxQuote.SuperTip = """Applies Box Quote style to selection"""
        '
        'grpBoxes_BoxQuoteListBullet
        '
        Me.grpBoxes_BoxQuoteListBullet.Label = "Box Quote List Bullet style"
        Me.grpBoxes_BoxQuoteListBullet.Name = "grpBoxes_BoxQuoteListBullet"
        Me.grpBoxes_BoxQuoteListBullet.OfficeImageId = "Q"
        Me.grpBoxes_BoxQuoteListBullet.ScreenTip = "Box Quote List Bullet style"
        Me.grpBoxes_BoxQuoteListBullet.ShowImage = True
        Me.grpBoxes_BoxQuoteListBullet.SuperTip = """Applies Box Quote List Bullet style to selection"""
        '
        'grpBoxes_BoxQuoteSource
        '
        Me.grpBoxes_BoxQuoteSource.Label = "Box Quote Source"
        Me.grpBoxes_BoxQuoteSource.Name = "grpBoxes_BoxQuoteSource"
        Me.grpBoxes_BoxQuoteSource.OfficeImageId = "Q"
        Me.grpBoxes_BoxQuoteSource.ScreenTip = "Box Quote Source"
        Me.grpBoxes_BoxQuoteSource.ShowImage = True
        Me.grpBoxes_BoxQuoteSource.SuperTip = """Applies Box Quote source style to selection"""
        '
        'Separator35
        '
        Me.Separator35.Name = "Separator35"
        '
        'grpBoxes_boxContent_mnu
        '
        Me.grpBoxes_boxContent_mnu.Items.Add(Me.grpBoxes_deleteBoxContent)
        Me.grpBoxes_boxContent_mnu.Items.Add(Me.grpBoxes_fillWithExampleStyles)
        Me.grpBoxes_boxContent_mnu.Label = "Box text for overtyping"
        Me.grpBoxes_boxContent_mnu.Name = "grpBoxes_boxContent_mnu"
        Me.grpBoxes_boxContent_mnu.OfficeImageId = "AppointmentColor4"
        Me.grpBoxes_boxContent_mnu.ScreenTip = "Box text for overtyping"
        Me.grpBoxes_boxContent_mnu.ShowImage = True
        Me.grpBoxes_boxContent_mnu.SuperTip = """This menu provides various default fill options for the selected Box"""
        '
        'grpBoxes_deleteBoxContent
        '
        Me.grpBoxes_deleteBoxContent.Label = "Delete and fill with 'OverType Here'Delete and fill with 'OverType Here'"
        Me.grpBoxes_deleteBoxContent.Name = "grpBoxes_deleteBoxContent"
        Me.grpBoxes_deleteBoxContent.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_deleteBoxContent.ShowImage = True
        Me.grpBoxes_deleteBoxContent.SuperTip = """Delete the contents and replace with 'OverType here'"""
        '
        'grpBoxes_fillWithExampleStyles
        '
        Me.grpBoxes_fillWithExampleStyles.Label = "Delete and fill with example styles"
        Me.grpBoxes_fillWithExampleStyles.Name = "grpBoxes_fillWithExampleStyles"
        Me.grpBoxes_fillWithExampleStyles.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_fillWithExampleStyles.ScreenTip = "Delete and fill with example styles"
        Me.grpBoxes_fillWithExampleStyles.ShowImage = True
        Me.grpBoxes_fillWithExampleStyles.SuperTip = """Delete the contents and replace with available/example styles"""
        '
        'mnuCloseDocuments1
        '
        Me.mnuCloseDocuments1.Items.Add(Me.grpBoxes_ToES)
        Me.mnuCloseDocuments1.Items.Add(Me.grpBoxes_ToBox1)
        Me.mnuCloseDocuments1.Items.Add(Me.grpBoxes_ToApp)
        Me.mnuCloseDocuments1.Items.Add(Me.Separator36)
        Me.mnuCloseDocuments1.Items.Add(Me.grpBoxes_ToLT)
        Me.mnuCloseDocuments1.KeyTip = "B#"
        Me.mnuCloseDocuments1.Label = "Convert Box #"
        Me.mnuCloseDocuments1.Name = "mnuCloseDocuments1"
        Me.mnuCloseDocuments1.OfficeImageId = "AppointmentColor4"
        Me.mnuCloseDocuments1.ScreenTip = "Convert Box #"
        Me.mnuCloseDocuments1.ShowImage = True
        Me.mnuCloseDocuments1.SuperTip = """Use functions from this menu to convert Box numbering to a different numbering f" &
    "ormat."""
        '
        'grpBoxes_ToES
        '
        Me.grpBoxes_ToES.Label = "Convert to Box (ES)"
        Me.grpBoxes_ToES.Name = "grpBoxes_ToES"
        Me.grpBoxes_ToES.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_ToES.ScreenTip = "Convert to Box (ES)"
        Me.grpBoxes_ToES.ShowImage = True
        Me.grpBoxes_ToES.SuperTip = """This menu item works on individual placeholder captions. Click in the Box captio" &
    "n, then click on this item to convert to an 'ES numbered box'."""
        '
        'grpBoxes_ToBox1
        '
        Me.grpBoxes_ToBox1.Label = "Convert to Box (Report)"
        Me.grpBoxes_ToBox1.Name = "grpBoxes_ToBox1"
        Me.grpBoxes_ToBox1.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_ToBox1.ScreenTip = "Convert to Box (Report)"
        Me.grpBoxes_ToBox1.ShowImage = True
        Me.grpBoxes_ToBox1.SuperTip = """This menu item works on individual placeholder captions. Click in the Box captio" &
    "n, then click on this item to convert to a 'Report numbered box'."""
        '
        'grpBoxes_ToApp
        '
        Me.grpBoxes_ToApp.Label = "Convert to Box (Appendix)"
        Me.grpBoxes_ToApp.Name = "grpBoxes_ToApp"
        Me.grpBoxes_ToApp.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_ToApp.ScreenTip = "Convert to Box (Appendix)"
        Me.grpBoxes_ToApp.ShowImage = True
        Me.grpBoxes_ToApp.SuperTip = """This menu item works on individual placeholder captions. Click in the Box headin" &
    "g, then click on this item to convert to 'Appendix numbering'."""
        '
        'Separator36
        '
        Me.Separator36.Name = "Separator36"
        '
        'grpBoxes_ToLT
        '
        Me.grpBoxes_ToLT.Label = "Convert to Box (Letter)"
        Me.grpBoxes_ToLT.Name = "grpBoxes_ToLT"
        Me.grpBoxes_ToLT.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_ToLT.ScreenTip = "Convert to Box (Letter)"
        Me.grpBoxes_ToLT.ShowImage = True
        Me.grpBoxes_ToLT.SuperTip = """This menu item works on individual placeholder captions. Click in the Box headin" &
    "g, then click on this item button to convert to 'Letter numbering'."""
        '
        'mnu_grpBoxes_Recommendations
        '
        Me.mnu_grpBoxes_Recommendations.Items.Add(Me.grpBoxes_Recommendation)
        Me.mnu_grpBoxes_Recommendations.Items.Add(Me.grpBoxes_RecommendationES)
        Me.mnu_grpBoxes_Recommendations.KeyTip = "FR"
        Me.mnu_grpBoxes_Recommendations.Label = "Recommendations"
        Me.mnu_grpBoxes_Recommendations.Name = "mnu_grpBoxes_Recommendations"
        Me.mnu_grpBoxes_Recommendations.OfficeImageId = "BevelShapeGallery"
        Me.mnu_grpBoxes_Recommendations.ScreenTip = "Recommendations"
        Me.mnu_grpBoxes_Recommendations.ShowImage = True
        '
        'grpBoxes_Recommendation
        '
        Me.grpBoxes_Recommendation.Label = "Recommendation (&Report)"
        Me.grpBoxes_Recommendation.Name = "grpBoxes_Recommendation"
        Me.grpBoxes_Recommendation.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_Recommendation.ScreenTip = "Recommendation (Report)"
        Me.grpBoxes_Recommendation.ShowImage = True
        Me.grpBoxes_Recommendation.SuperTip = """Inserts a report 'Recommendations' box that spans from the left edge of the page" &
    " to the right edge."""
        '
        'grpBoxes_RecommendationES
        '
        Me.grpBoxes_RecommendationES.Label = "Recommendation (&ES)"
        Me.grpBoxes_RecommendationES.Name = "grpBoxes_RecommendationES"
        Me.grpBoxes_RecommendationES.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_RecommendationES.ScreenTip = "Recommendation (ES)"
        Me.grpBoxes_RecommendationES.ShowImage = True
        Me.grpBoxes_RecommendationES.SuperTip = """Inserts an es 'Recommendations' box that spans from the left edge of the page to" &
    " the right edge."""
        '
        'mnu_grpBoxes_Findings
        '
        Me.mnu_grpBoxes_Findings.Items.Add(Me.grpBoxes_KeyFinding)
        Me.mnu_grpBoxes_Findings.Items.Add(Me.grpBoxes_KeyFindingES)
        Me.mnu_grpBoxes_Findings.KeyTip = "FI"
        Me.mnu_grpBoxes_Findings.Label = "Findings"
        Me.mnu_grpBoxes_Findings.Name = "mnu_grpBoxes_Findings"
        Me.mnu_grpBoxes_Findings.OfficeImageId = "BevelShapeGallery"
        Me.mnu_grpBoxes_Findings.ScreenTip = "Findings"
        Me.mnu_grpBoxes_Findings.ShowImage = True
        '
        'grpBoxes_KeyFinding
        '
        Me.grpBoxes_KeyFinding.Label = "Finding (&Report)"
        Me.grpBoxes_KeyFinding.Name = "grpBoxes_KeyFinding"
        Me.grpBoxes_KeyFinding.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_KeyFinding.ScreenTip = "Finding (Report)"
        Me.grpBoxes_KeyFinding.ShowImage = True
        Me.grpBoxes_KeyFinding.SuperTip = """Inserts a report 'Findings' box that spans from the left edge of the page to the" &
    " right edge."""
        '
        'grpBoxes_KeyFindingES
        '
        Me.grpBoxes_KeyFindingES.Label = "Finding (&ES)"
        Me.grpBoxes_KeyFindingES.Name = "grpBoxes_KeyFindingES"
        Me.grpBoxes_KeyFindingES.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_KeyFindingES.ScreenTip = "Finding (ES)"
        Me.grpBoxes_KeyFindingES.ShowImage = True
        Me.grpBoxes_KeyFindingES.SuperTip = """Inserts an es 'Findings' box that spans from the left edge of the page to the ri" &
    "ght edge."""
        '
        'grpPullouts_mnu01
        '
        Me.grpPullouts_mnu01.Items.Add(Me.grpPullouts_emphasisBox_Left)
        Me.grpPullouts_mnu01.Items.Add(Me.grpPullouts_emphasisBox_Centre)
        Me.grpPullouts_mnu01.Items.Add(Me.grpPullouts_emphasisBox_Right)
        Me.grpPullouts_mnu01.KeyTip = "MP"
        Me.grpPullouts_mnu01.Label = "Emphasis"
        Me.grpPullouts_mnu01.Name = "grpPullouts_mnu01"
        Me.grpPullouts_mnu01.OfficeImageId = "ControlLogo"
        Me.grpPullouts_mnu01.ScreenTip = "Emphasis"
        Me.grpPullouts_mnu01.ShowImage = True
        Me.grpPullouts_mnu01.SuperTip = """Select various Emphasis boxes to highlight important issues. Use the 'Emphasis' " &
    "text styles in documents that need to be 'Accessible' (not the boxes)."""
        '
        'grpPullouts_emphasisBox_Left
        '
        Me.grpPullouts_emphasisBox_Left.Label = "Emphasis Box (&left)"
        Me.grpPullouts_emphasisBox_Left.Name = "grpPullouts_emphasisBox_Left"
        Me.grpPullouts_emphasisBox_Left.OfficeImageId = "ControlLogo"
        Me.grpPullouts_emphasisBox_Left.ScreenTip = "Emphasis Box (&left)"
        Me.grpPullouts_emphasisBox_Left.ShowImage = True
        Me.grpPullouts_emphasisBox_Left.SuperTip = """Inserts an emphasis box aligned to the left hand margin (text is left justified)" &
    ". Do not use this element in documents that need to be 'Accessible'"""
        '
        'grpPullouts_emphasisBox_Centre
        '
        Me.grpPullouts_emphasisBox_Centre.Label = "Emphasis Box (&centre)"
        Me.grpPullouts_emphasisBox_Centre.Name = "grpPullouts_emphasisBox_Centre"
        Me.grpPullouts_emphasisBox_Centre.OfficeImageId = "ControlLogo"
        Me.grpPullouts_emphasisBox_Centre.ScreenTip = "Emphasis Box (centre)"
        Me.grpPullouts_emphasisBox_Centre.ShowImage = True
        Me.grpPullouts_emphasisBox_Centre.SuperTip = """Inserts an emphasis box centred between the left and right margins (text is cent" &
    "re justified). Do not use this element in documents that need to be 'Accessible'" &
    """"
        '
        'grpPullouts_emphasisBox_Right
        '
        Me.grpPullouts_emphasisBox_Right.Label = "Emphasis Box (&right)"
        Me.grpPullouts_emphasisBox_Right.Name = "grpPullouts_emphasisBox_Right"
        Me.grpPullouts_emphasisBox_Right.OfficeImageId = "ControlLogo"
        Me.grpPullouts_emphasisBox_Right.ScreenTip = "Emphasis Box (right)"
        Me.grpPullouts_emphasisBox_Right.ShowImage = True
        Me.grpPullouts_emphasisBox_Right.SuperTip = """Inserts an emphasis box aligned to the right hand margin (text is right justifie" &
    "d). Do not use this element in documents that need to be 'Accessible'"""
        '
        'grpReport_mnu_CaseStudies
        '
        Me.grpReport_mnu_CaseStudies.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grpReport_mnu_CaseStudies.Items.Add(Me.grpReport_mnu_CaseStudies_FullPage)
        Me.grpReport_mnu_CaseStudies.Items.Add(Me.grpReport_mnu_CaseStudies_HalfPage)
        Me.grpReport_mnu_CaseStudies.Items.Add(Me.Separator39)
        Me.grpReport_mnu_CaseStudies.Items.Add(Me.grpReport_mnu_CaseStudies_CaseStudyHeading)
        Me.grpReport_mnu_CaseStudies.KeyTip = "RS"
        Me.grpReport_mnu_CaseStudies.Label = "Case Studies"
        Me.grpReport_mnu_CaseStudies.Name = "grpReport_mnu_CaseStudies"
        Me.grpReport_mnu_CaseStudies.OfficeImageId = "TextBoxInsert"
        Me.grpReport_mnu_CaseStudies.ScreenTip = "Case Studies"
        Me.grpReport_mnu_CaseStudies.ShowImage = True
        Me.grpReport_mnu_CaseStudies.SuperTip = """This menu provides functions that allow you to insert Case Study segments to the" &
    " document. Single page, multiple pages or part of a page"""
        '
        'grpReport_mnu_CaseStudies_FullPage
        '
        Me.grpReport_mnu_CaseStudies_FullPage.Label = "&Full Page Case Study Placeholder"
        Me.grpReport_mnu_CaseStudies_FullPage.Name = "grpReport_mnu_CaseStudies_FullPage"
        Me.grpReport_mnu_CaseStudies_FullPage.OfficeImageId = "TextBoxInsert"
        Me.grpReport_mnu_CaseStudies_FullPage.ScreenTip = "Full Page Case Study Placeholder"
        Me.grpReport_mnu_CaseStudies_FullPage.ShowImage = True
        Me.grpReport_mnu_CaseStudies_FullPage.SuperTip = """Inserts a new section allowing for a single or multiple page Case Study"""
        '
        'grpReport_mnu_CaseStudies_HalfPage
        '
        Me.grpReport_mnu_CaseStudies_HalfPage.Label = "&Part Page Case Study Placeholder"
        Me.grpReport_mnu_CaseStudies_HalfPage.Name = "grpReport_mnu_CaseStudies_HalfPage"
        Me.grpReport_mnu_CaseStudies_HalfPage.OfficeImageId = "TextBoxInsert"
        Me.grpReport_mnu_CaseStudies_HalfPage.ScreenTip = "Part Page Case Study Placeholder"
        Me.grpReport_mnu_CaseStudies_HalfPage.ShowImage = True
        Me.grpReport_mnu_CaseStudies_HalfPage.SuperTip = """Inserts a new placeholder (typically half a page) for small Case Studies"""
        '
        'Separator39
        '
        Me.Separator39.Name = "Separator39"
        '
        'grpReport_mnu_CaseStudies_CaseStudyHeading
        '
        Me.grpReport_mnu_CaseStudies_CaseStudyHeading.Label = "&Heading (CaseStudy numbered caption)"
        Me.grpReport_mnu_CaseStudies_CaseStudyHeading.Name = "grpReport_mnu_CaseStudies_CaseStudyHeading"
        Me.grpReport_mnu_CaseStudies_CaseStudyHeading.OfficeImageId = "TextBoxInsert"
        Me.grpReport_mnu_CaseStudies_CaseStudyHeading.ScreenTip = "Heading (CaseStudy numbered caption)"
        Me.grpReport_mnu_CaseStudies_CaseStudyHeading.ShowImage = True
        Me.grpReport_mnu_CaseStudies_CaseStudyHeading.SuperTip = """Will apply the case study numbered caption heading to the selected paragraphs"""
        '
        'Separator37
        '
        Me.Separator37.Name = "Separator37"
        '
        'mnuCloseDocuments16
        '
        Me.mnuCloseDocuments16.Items.Add(Me.grpFigures_Figure)
        Me.mnuCloseDocuments16.Items.Add(Me.Separator44)
        Me.mnuCloseDocuments16.Items.Add(Me.grpFigures_Appendix)
        Me.mnuCloseDocuments16.Items.Add(Me.Separator43)
        Me.mnuCloseDocuments16.Items.Add(Me.grpFigures_ES)
        Me.mnuCloseDocuments16.Items.Add(Me.Separator42)
        Me.mnuCloseDocuments16.Items.Add(Me.grpFigures_LT)
        Me.mnuCloseDocuments16.Items.Add(Me.Separator41)
        Me.mnuCloseDocuments16.Items.Add(Me.grpFigures_CaptionAndHeading)
        Me.mnuCloseDocuments16.Items.Add(Me.grpFigures_CaptionAndHeadingApp)
        Me.mnuCloseDocuments16.Items.Add(Me.grpFigures_CaptionAndHeadingES)
        Me.mnuCloseDocuments16.Items.Add(Me.Separator40)
        Me.mnuCloseDocuments16.Items.Add(Me.grpFigures_StyleForSubHeadings)
        Me.mnuCloseDocuments16.KeyTip = "FG"
        Me.mnuCloseDocuments16.Label = "Figures"
        Me.mnuCloseDocuments16.Name = "mnuCloseDocuments16"
        Me.mnuCloseDocuments16.OfficeImageId = "ControlLogo"
        Me.mnuCloseDocuments16.ScreenTip = "Figures"
        Me.mnuCloseDocuments16.ShowImage = True
        Me.mnuCloseDocuments16.SuperTip = """Select from the list of preformatted sized placeholders which have the caption l" &
    "abel Figure."""
        '
        'grpFigures_Figure
        '
        Me.grpFigures_Figure.Label = "&Figure"
        Me.grpFigures_Figure.Name = "grpFigures_Figure"
        Me.grpFigures_Figure.OfficeImageId = "ControlLogo"
        Me.grpFigures_Figure.ScreenTip = "Figure"
        Me.grpFigures_Figure.ShowImage = True
        Me.grpFigures_Figure.SuperTip = """Inserts a placeholder with Figure caption"""
        '
        'Separator44
        '
        Me.Separator44.Name = "Separator44"
        '
        'grpFigures_Appendix
        '
        Me.grpFigures_Appendix.Label = "&Appendix Figure"
        Me.grpFigures_Appendix.Name = "grpFigures_Appendix"
        Me.grpFigures_Appendix.OfficeImageId = "ControlLogo"
        Me.grpFigures_Appendix.ScreenTip = "Appendix Figure"
        Me.grpFigures_Appendix.ShowImage = True
        Me.grpFigures_Appendix.SuperTip = """Inserts a placeholder to fit within text margins with Figure caption prefixed by" &
    " Appendix number"""
        '
        'Separator43
        '
        Me.Separator43.Name = "Separator43"
        '
        'grpFigures_ES
        '
        Me.grpFigures_ES.Label = "Figure &ES"
        Me.grpFigures_ES.Name = "grpFigures_ES"
        Me.grpFigures_ES.OfficeImageId = "ControlLogo"
        Me.grpFigures_ES.ScreenTip = "Figure ES"
        Me.grpFigures_ES.ShowImage = True
        Me.grpFigures_ES.SuperTip = """Inserts a placeholder with Figure ES caption to fit within text margin."""
        '
        'Separator42
        '
        Me.Separator42.Name = "Separator42"
        '
        'grpFigures_LT
        '
        Me.grpFigures_LT.Label = "Figure &LT (for letters, memos etc)"
        Me.grpFigures_LT.Name = "grpFigures_LT"
        Me.grpFigures_LT.OfficeImageId = "ControlLogo"
        Me.grpFigures_LT.ScreenTip = "Figure &LT (for letters, memos etc)"
        Me.grpFigures_LT.ShowImage = True
        Me.grpFigures_LT.SuperTip = """To be used for letters, memos etc. This menu item inserts a Figure placeholder t" &
    "hat uses 'Stationery Numbering' for the Caption."""
        '
        'Separator41
        '
        Me.Separator41.Name = "Separator41"
        '
        'grpFigures_CaptionAndHeading
        '
        Me.grpFigures_CaptionAndHeading.Label = "Insert Figure &Caption and Heading only"
        Me.grpFigures_CaptionAndHeading.Name = "grpFigures_CaptionAndHeading"
        Me.grpFigures_CaptionAndHeading.OfficeImageId = "ControlLogo"
        Me.grpFigures_CaptionAndHeading.ScreenTip = "Figure Caption and Heading"
        Me.grpFigures_CaptionAndHeading.ShowImage = True
        '
        'grpFigures_CaptionAndHeadingApp
        '
        Me.grpFigures_CaptionAndHeadingApp.Label = "Insert Appendix Figure Caption and Heading only"
        Me.grpFigures_CaptionAndHeadingApp.Name = "grpFigures_CaptionAndHeadingApp"
        Me.grpFigures_CaptionAndHeadingApp.OfficeImageId = "ControlLogo"
        Me.grpFigures_CaptionAndHeadingApp.ScreenTip = "Appendix Figure Caption and Heading"
        Me.grpFigures_CaptionAndHeadingApp.ShowImage = True
        Me.grpFigures_CaptionAndHeadingApp.SuperTip = """Inserts Appendix numbered Figure Caption and number only. """
        '
        'grpFigures_CaptionAndHeadingES
        '
        Me.grpFigures_CaptionAndHeadingES.Label = "Insert ES Figure Caption and Heading only"
        Me.grpFigures_CaptionAndHeadingES.Name = "grpFigures_CaptionAndHeadingES"
        Me.grpFigures_CaptionAndHeadingES.OfficeImageId = "ControlLogo"
        Me.grpFigures_CaptionAndHeadingES.ScreenTip = "ES Figure Caption and Heading"
        Me.grpFigures_CaptionAndHeadingES.ShowImage = True
        Me.grpFigures_CaptionAndHeadingES.SuperTip = """Inserts ES prefixed Figure Caption and number only. """
        '
        'Separator40
        '
        Me.Separator40.Name = "Separator40"
        '
        'grpFigures_StyleForSubHeadings
        '
        Me.grpFigures_StyleForSubHeadings.Label = "Figure style for sub headings within placeholder"
        Me.grpFigures_StyleForSubHeadings.Name = "grpFigures_StyleForSubHeadings"
        Me.grpFigures_StyleForSubHeadings.OfficeImageId = "DataOptionsMenu"
        Me.grpFigures_StyleForSubHeadings.ScreenTip = "Figure style for sub headings within placeholder"
        Me.grpFigures_StyleForSubHeadings.ShowImage = True
        Me.grpFigures_StyleForSubHeadings.SuperTip = resources.GetString("grpFigures_StyleForSubHeadings.SuperTip")
        '
        'mnuCloseDocuments33
        '
        Me.mnuCloseDocuments33.Items.Add(Me.grpFigures_convertToES)
        Me.mnuCloseDocuments33.Items.Add(Me.grpFigures_convertToStd)
        Me.mnuCloseDocuments33.Items.Add(Me.grpFigures_convertToApp)
        Me.mnuCloseDocuments33.Items.Add(Me.Separator45)
        Me.mnuCloseDocuments33.Items.Add(Me.grpFigures_convertToLT)
        Me.mnuCloseDocuments33.KeyTip = "CF"
        Me.mnuCloseDocuments33.Label = "Convert Figure #"
        Me.mnuCloseDocuments33.Name = "mnuCloseDocuments33"
        Me.mnuCloseDocuments33.OfficeImageId = "OmsChangeSlideLayoutGallery"
        Me.mnuCloseDocuments33.ScreenTip = "Convert Figure #"
        Me.mnuCloseDocuments33.ShowImage = True
        Me.mnuCloseDocuments33.SuperTip = """Use items from this menu to convert Figure numbering to a different numbering fo" &
    "rmat."""
        '
        'grpFigures_convertToES
        '
        Me.grpFigures_convertToES.Label = "Convert to Figure (&ES)"
        Me.grpFigures_convertToES.Name = "grpFigures_convertToES"
        Me.grpFigures_convertToES.OfficeImageId = "OmsChangeSlideLayoutGallery"
        Me.grpFigures_convertToES.ScreenTip = "Convert to Figure (ES)"
        Me.grpFigures_convertToES.ShowImage = True
        Me.grpFigures_convertToES.SuperTip = """This menu item works on individual placeholder captions. Click in the Figure cap" &
    "tion, then click on this item button to convert to 'ES numbering'."""
        '
        'grpFigures_convertToStd
        '
        Me.grpFigures_convertToStd.Label = "Convert to Figure (&Report)"
        Me.grpFigures_convertToStd.Name = "grpFigures_convertToStd"
        Me.grpFigures_convertToStd.OfficeImageId = "OmsChangeSlideLayoutGallery"
        Me.grpFigures_convertToStd.ScreenTip = "Convert to Figure (Report)"
        Me.grpFigures_convertToStd.ShowImage = True
        Me.grpFigures_convertToStd.SuperTip = """This menu item works on individual placeholder captions. Click in the Figure cap" &
    "tion, then click on this item button to convert to 'Report numbering'."""
        '
        'grpFigures_convertToApp
        '
        Me.grpFigures_convertToApp.Label = "Convert to Figure (&Appendix)"
        Me.grpFigures_convertToApp.Name = "grpFigures_convertToApp"
        Me.grpFigures_convertToApp.OfficeImageId = "OmsChangeSlideLayoutGallery"
        Me.grpFigures_convertToApp.ScreenTip = "Convert to Figure (Appendix)"
        Me.grpFigures_convertToApp.ShowImage = True
        Me.grpFigures_convertToApp.SuperTip = """This menu item works on individual placeholder captions. Click in the Figure cap" &
    "tion, then click on this item button to convert to 'Appendix numbering'."""
        '
        'Separator45
        '
        Me.Separator45.Name = "Separator45"
        '
        'grpFigures_convertToLT
        '
        Me.grpFigures_convertToLT.Label = "Convert to Figure (&Letter)"
        Me.grpFigures_convertToLT.Name = "grpFigures_convertToLT"
        Me.grpFigures_convertToLT.OfficeImageId = "OmsChangeSlideLayoutGallery"
        Me.grpFigures_convertToLT.ScreenTip = "Convert to Figure (Letter)"
        Me.grpFigures_convertToLT.ShowImage = True
        Me.grpFigures_convertToLT.SuperTip = """This menu item works on individual placeholder captions. Click in the Figure cap" &
    "tion, then click on this item button to convert to 'Letter numbering'."""
        '
        'grpStylesRpt_mnu_tbls_00
        '
        Me.grpStylesRpt_mnu_tbls_00.Items.Add(Me.grpTbls_fillCellsWithCustomColour)
        Me.grpStylesRpt_mnu_tbls_00.Items.Add(Me.Separator46)
        Me.grpStylesRpt_mnu_tbls_00.Items.Add(Me.grpTbls_setTableTextCustomColour)
        Me.grpStylesRpt_mnu_tbls_00.Label = "Apply Colour"
        Me.grpStylesRpt_mnu_tbls_00.Name = "grpStylesRpt_mnu_tbls_00"
        Me.grpStylesRpt_mnu_tbls_00.OfficeImageId = "ViewBackToColorView"
        Me.grpStylesRpt_mnu_tbls_00.ScreenTip = "Apply Colour"
        Me.grpStylesRpt_mnu_tbls_00.ShowImage = True
        Me.grpStylesRpt_mnu_tbls_00.SuperTip = """Apply a specific colour to the currently selected table cells"""
        '
        'grpTbls_fillCellsWithCustomColour
        '
        Me.grpTbls_fillCellsWithCustomColour.Label = "Apply to selected table cells"
        Me.grpTbls_fillCellsWithCustomColour.Name = "grpTbls_fillCellsWithCustomColour"
        Me.grpTbls_fillCellsWithCustomColour.OfficeImageId = "ViewBackToColorView"
        Me.grpTbls_fillCellsWithCustomColour.ScreenTip = "Apply to selected table cells"
        Me.grpTbls_fillCellsWithCustomColour.ShowImage = True
        Me.grpTbls_fillCellsWithCustomColour.SuperTip = """Shows a dialog that allows the user to apply a custom colour to the selected tab" &
    "le cells """
        '
        'Separator46
        '
        Me.Separator46.Name = "Separator46"
        '
        'grpTbls_setTableTextCustomColour
        '
        Me.grpTbls_setTableTextCustomColour.Label = "Apply to selected text"
        Me.grpTbls_setTableTextCustomColour.Name = "grpTbls_setTableTextCustomColour"
        Me.grpTbls_setTableTextCustomColour.OfficeImageId = "ViewBackToColorView"
        Me.grpTbls_setTableTextCustomColour.ScreenTip = "Apply to selected text"
        Me.grpTbls_setTableTextCustomColour.ShowImage = True
        Me.grpTbls_setTableTextCustomColour.SuperTip = """Shows a dialog that allows the user to apply a custom colour to the selected tex" &
    "t (table text or general body text)."""
        '
        'Separator38
        '
        Me.Separator38.Name = "Separator38"
        '
        'grpPlh_btn_buildCustomTable
        '
        Me.grpPlh_btn_buildCustomTable.KeyTip = "TC"
        Me.grpPlh_btn_buildCustomTable.Label = "Custom Table"
        Me.grpPlh_btn_buildCustomTable.Name = "grpPlh_btn_buildCustomTable"
        Me.grpPlh_btn_buildCustomTable.OfficeImageId = "FieldChooser"
        Me.grpPlh_btn_buildCustomTable.ShowImage = True
        Me.grpPlh_btn_buildCustomTable.SuperTip = resources.GetString("grpPlh_btn_buildCustomTable.SuperTip")
        '
        'grpTbl_Styles
        '
        Me.grpTbl_Styles.Items.Add(Me.grpTbls_TableColumnHeadingsStyle)
        Me.grpTbl_Styles.Items.Add(Me.grpTbls_TableUnitsRowStyle)
        Me.grpTbl_Styles.Items.Add(Me.Separator47)
        Me.grpTbl_Styles.Items.Add(Me.grpTbls_Plh_mnu_TableListBulletsStyles)
        Me.grpTbl_Styles.Items.Add(Me.grpTbls_Plh_mnu_TableListNumberingStyles)
        Me.grpTbl_Styles.Items.Add(Me.Separator48)
        Me.grpTbl_Styles.Items.Add(Me.grpTbls_Plh_mnu_SideHeadingStyles)
        Me.grpTbl_Styles.Items.Add(Me.grpTbls_Plh_mnu_QuoteStyles)
        Me.grpTbl_Styles.Items.Add(Me.Separator49)
        Me.grpTbl_Styles.Items.Add(Me.grpTbl_Styles_ExampleStyleSets)
        Me.grpTbl_Styles.Items.Add(Me.Separator50)
        Me.grpTbl_Styles.Items.Add(Me.grpTbls_ColourCells)
        Me.grpTbl_Styles.Items.Add(Me.grpTbls_ColourHeadingsRow)
        Me.grpTbl_Styles.Items.Add(Me.grpTbls_ColourUnitsRow)
        Me.grpTbl_Styles.Items.Add(Me.Separator51)
        Me.grpTbl_Styles.Items.Add(Me.grpTbls_AllBorders)
        Me.grpTbl_Styles.Items.Add(Me.grpTbls_AllBordersRemove)
        Me.grpTbl_Styles.KeyTip = "TS"
        Me.grpTbl_Styles.Label = "Table Styles"
        Me.grpTbl_Styles.Name = "grpTbl_Styles"
        Me.grpTbl_Styles.OfficeImageId = "BevelShapeGallery"
        Me.grpTbl_Styles.ScreenTip = "Table Styles"
        Me.grpTbl_Styles.ShowImage = True
        Me.grpTbl_Styles.SuperTip = """These menu items allow you to manipulate various styles and aspects of the selec" &
    "ted table."""
        '
        'grpTbls_TableColumnHeadingsStyle
        '
        Me.grpTbls_TableColumnHeadingsStyle.Label = "Table Column &Headings style"
        Me.grpTbls_TableColumnHeadingsStyle.Name = "grpTbls_TableColumnHeadingsStyle"
        Me.grpTbls_TableColumnHeadingsStyle.OfficeImageId = "AccessTableEvents"
        Me.grpTbls_TableColumnHeadingsStyle.ScreenTip = "Table Column Headings style"
        Me.grpTbls_TableColumnHeadingsStyle.ShowImage = True
        Me.grpTbls_TableColumnHeadingsStyle.SuperTip = """Applies Table Column Headings style to selection"""
        '
        'grpTbls_TableUnitsRowStyle
        '
        Me.grpTbls_TableUnitsRowStyle.Label = "Table &Units Row style"
        Me.grpTbls_TableUnitsRowStyle.Name = "grpTbls_TableUnitsRowStyle"
        Me.grpTbls_TableUnitsRowStyle.OfficeImageId = "RecordsMoreRecordsMenu"
        Me.grpTbls_TableUnitsRowStyle.ScreenTip = "Table Units Row style"
        Me.grpTbls_TableUnitsRowStyle.ShowImage = True
        Me.grpTbls_TableUnitsRowStyle.SuperTip = """Applies Table Units Row style to selection"""
        '
        'Separator47
        '
        Me.Separator47.Name = "Separator47"
        '
        'grpTbls_Plh_mnu_TableListBulletsStyles
        '
        Me.grpTbls_Plh_mnu_TableListBulletsStyles.Items.Add(Me.grpTbls_TableListBullet)
        Me.grpTbls_Plh_mnu_TableListBulletsStyles.Items.Add(Me.grpTbls_TableListBullet2)
        Me.grpTbls_Plh_mnu_TableListBulletsStyles.Items.Add(Me.grpTbls_TableListBullet3)
        Me.grpTbls_Plh_mnu_TableListBulletsStyles.Label = "Table List &Bullet Styles"
        Me.grpTbls_Plh_mnu_TableListBulletsStyles.Name = "grpTbls_Plh_mnu_TableListBulletsStyles"
        Me.grpTbls_Plh_mnu_TableListBulletsStyles.OfficeImageId = "Bullets"
        Me.grpTbls_Plh_mnu_TableListBulletsStyles.ScreenTip = "Standard Table Styles"
        Me.grpTbls_Plh_mnu_TableListBulletsStyles.ShowImage = True
        '
        'grpTbls_TableListBullet
        '
        Me.grpTbls_TableListBullet.Label = "Table List Bullet &1 style"
        Me.grpTbls_TableListBullet.Name = "grpTbls_TableListBullet"
        Me.grpTbls_TableListBullet.OfficeImageId = "Bullets"
        Me.grpTbls_TableListBullet.ScreenTip = "Standard List Bullet"
        Me.grpTbls_TableListBullet.ShowImage = True
        Me.grpTbls_TableListBullet.SuperTip = """Applies Table List Bullet style to selection"""
        '
        'grpTbls_TableListBullet2
        '
        Me.grpTbls_TableListBullet2.Label = "Table List Bullet &2 style"
        Me.grpTbls_TableListBullet2.Name = "grpTbls_TableListBullet2"
        Me.grpTbls_TableListBullet2.OfficeImageId = "Bullets"
        Me.grpTbls_TableListBullet2.ScreenTip = "Standard List Bullet"
        Me.grpTbls_TableListBullet2.ShowImage = True
        Me.grpTbls_TableListBullet2.SuperTip = "Applies Table List Bullet 2 style to selection"
        '
        'grpTbls_TableListBullet3
        '
        Me.grpTbls_TableListBullet3.Label = "Table List Bullet &3 style"
        Me.grpTbls_TableListBullet3.Name = "grpTbls_TableListBullet3"
        Me.grpTbls_TableListBullet3.OfficeImageId = "Bullets"
        Me.grpTbls_TableListBullet3.ScreenTip = "Standard List Bullet"
        Me.grpTbls_TableListBullet3.ShowImage = True
        Me.grpTbls_TableListBullet3.SuperTip = """Applies Table List Bullet 3 style to selection"""
        '
        'grpTbls_Plh_mnu_TableListNumberingStyles
        '
        Me.grpTbls_Plh_mnu_TableListNumberingStyles.Items.Add(Me.grpTbls_ListNumber)
        Me.grpTbls_Plh_mnu_TableListNumberingStyles.Items.Add(Me.grpTbls_ListNumber2)
        Me.grpTbls_Plh_mnu_TableListNumberingStyles.Items.Add(Me.grpTbls_ListNumber3)
        Me.grpTbls_Plh_mnu_TableListNumberingStyles.Label = "Table List &Number Styles"
        Me.grpTbls_Plh_mnu_TableListNumberingStyles.Name = "grpTbls_Plh_mnu_TableListNumberingStyles"
        Me.grpTbls_Plh_mnu_TableListNumberingStyles.OfficeImageId = "Numbering"
        Me.grpTbls_Plh_mnu_TableListNumberingStyles.ShowImage = True
        Me.grpTbls_Plh_mnu_TableListNumberingStyles.SuperTip = "Standard Table Styles"
        '
        'grpTbls_ListNumber
        '
        Me.grpTbls_ListNumber.Label = "Table List Number &1 style"
        Me.grpTbls_ListNumber.Name = "grpTbls_ListNumber"
        Me.grpTbls_ListNumber.OfficeImageId = "Numbering"
        Me.grpTbls_ListNumber.ScreenTip = "Standard List Number"
        Me.grpTbls_ListNumber.ShowImage = True
        Me.grpTbls_ListNumber.SuperTip = """Applies Table List Number style to selection"""
        '
        'grpTbls_ListNumber2
        '
        Me.grpTbls_ListNumber2.Label = "Table List Number &2 style"
        Me.grpTbls_ListNumber2.Name = "grpTbls_ListNumber2"
        Me.grpTbls_ListNumber2.OfficeImageId = "Numbering"
        Me.grpTbls_ListNumber2.ScreenTip = "Standard List Number"
        Me.grpTbls_ListNumber2.ShowImage = True
        Me.grpTbls_ListNumber2.SuperTip = """Applies Table List Number 2 style to selection"""
        '
        'grpTbls_ListNumber3
        '
        Me.grpTbls_ListNumber3.Label = "Table List Number &3 style"
        Me.grpTbls_ListNumber3.Name = "grpTbls_ListNumber3"
        Me.grpTbls_ListNumber3.OfficeImageId = "Numbering"
        Me.grpTbls_ListNumber3.ScreenTip = "Standard List Number"
        Me.grpTbls_ListNumber3.ShowImage = True
        Me.grpTbls_ListNumber3.SuperTip = """Applies Table List Number 3 style to selection"""
        '
        'Separator48
        '
        Me.Separator48.Name = "Separator48"
        '
        'grpTbls_Plh_mnu_SideHeadingStyles
        '
        Me.grpTbls_Plh_mnu_SideHeadingStyles.Items.Add(Me.grpTbls_TableSideHeading1)
        Me.grpTbls_Plh_mnu_SideHeadingStyles.Items.Add(Me.grpTbls_TableSideHeading2)
        Me.grpTbls_Plh_mnu_SideHeadingStyles.Label = "Table &Side Heading Styles"
        Me.grpTbls_Plh_mnu_SideHeadingStyles.Name = "grpTbls_Plh_mnu_SideHeadingStyles"
        Me.grpTbls_Plh_mnu_SideHeadingStyles.OfficeImageId = "BevelShapeGallery"
        Me.grpTbls_Plh_mnu_SideHeadingStyles.ShowImage = True
        '
        'grpTbls_TableSideHeading1
        '
        Me.grpTbls_TableSideHeading1.Label = "Table Side Heading &1"
        Me.grpTbls_TableSideHeading1.Name = "grpTbls_TableSideHeading1"
        Me.grpTbls_TableSideHeading1.OfficeImageId = "PivotCollapseFieldAccess"
        Me.grpTbls_TableSideHeading1.ScreenTip = "Table Side Heading 1"
        Me.grpTbls_TableSideHeading1.ShowImage = True
        Me.grpTbls_TableSideHeading1.SuperTip = """Applies Table Side Heading 1 style to selection"""
        '
        'grpTbls_TableSideHeading2
        '
        Me.grpTbls_TableSideHeading2.Label = "Table Side Heading &2"
        Me.grpTbls_TableSideHeading2.Name = "grpTbls_TableSideHeading2"
        Me.grpTbls_TableSideHeading2.ScreenTip = "Table Side Heading 2"
        Me.grpTbls_TableSideHeading2.ShowImage = True
        Me.grpTbls_TableSideHeading2.SuperTip = """Applies Table Side Heading 2 style to selection"""
        '
        'grpTbls_Plh_mnu_QuoteStyles
        '
        Me.grpTbls_Plh_mnu_QuoteStyles.Items.Add(Me.grpTbls_Quote)
        Me.grpTbls_Plh_mnu_QuoteStyles.Items.Add(Me.grpTbls_QuoteListBullet)
        Me.grpTbls_Plh_mnu_QuoteStyles.Items.Add(Me.grpTbls_QuoteSource)
        Me.grpTbls_Plh_mnu_QuoteStyles.Label = "Table &Quote Styles"
        Me.grpTbls_Plh_mnu_QuoteStyles.Name = "grpTbls_Plh_mnu_QuoteStyles"
        Me.grpTbls_Plh_mnu_QuoteStyles.OfficeImageId = "BevelShapeGallery"
        Me.grpTbls_Plh_mnu_QuoteStyles.ShowImage = True
        '
        'grpTbls_Quote
        '
        Me.grpTbls_Quote.Label = "Table &Quote style"
        Me.grpTbls_Quote.Name = "grpTbls_Quote"
        Me.grpTbls_Quote.OfficeImageId = "Q"
        Me.grpTbls_Quote.ScreenTip = "Table Quote style"
        Me.grpTbls_Quote.ShowImage = True
        Me.grpTbls_Quote.SuperTip = """Applies Table Quote style to selection"""
        '
        'grpTbls_QuoteListBullet
        '
        Me.grpTbls_QuoteListBullet.Label = "Table Quote List &Bullet style"
        Me.grpTbls_QuoteListBullet.Name = "grpTbls_QuoteListBullet"
        Me.grpTbls_QuoteListBullet.OfficeImageId = "Q"
        Me.grpTbls_QuoteListBullet.ScreenTip = "Table Quote List Bullet style"
        Me.grpTbls_QuoteListBullet.ShowImage = True
        Me.grpTbls_QuoteListBullet.SuperTip = """Applies Table Quote List Bullet style to selection"""
        '
        'grpTbls_QuoteSource
        '
        Me.grpTbls_QuoteSource.Label = "Table Quote &Source style"
        Me.grpTbls_QuoteSource.Name = "grpTbls_QuoteSource"
        Me.grpTbls_QuoteSource.OfficeImageId = "Q"
        Me.grpTbls_QuoteSource.ScreenTip = "Table Quote Source style"
        Me.grpTbls_QuoteSource.ShowImage = True
        Me.grpTbls_QuoteSource.SuperTip = """Applies Table Quote Source style to selection"""
        '
        'Separator49
        '
        Me.Separator49.Name = "Separator49"
        '
        'grpTbl_Styles_ExampleStyleSets
        '
        Me.grpTbl_Styles_ExampleStyleSets.Items.Add(Me.grpTbls_StyleSet_TableQuote)
        Me.grpTbl_Styles_ExampleStyleSets.Items.Add(Me.grpTbls_StyleSet_TableListBullets)
        Me.grpTbl_Styles_ExampleStyleSets.Items.Add(Me.grpTbls_StyleSet_TableListNumbers)
        Me.grpTbl_Styles_ExampleStyleSets.Label = "E&xample Style Sets (normal size)"
        Me.grpTbl_Styles_ExampleStyleSets.Name = "grpTbl_Styles_ExampleStyleSets"
        Me.grpTbl_Styles_ExampleStyleSets.OfficeImageId = "FunctionsLogicalInsertGallery"
        Me.grpTbl_Styles_ExampleStyleSets.ScreenTip = "Example Style Sets (normal size)"
        Me.grpTbl_Styles_ExampleStyleSets.ShowImage = True
        Me.grpTbl_Styles_ExampleStyleSets.SuperTip = """This menu will allow you to insert example style sets at the cursor position."""
        '
        'grpTbls_StyleSet_TableQuote
        '
        Me.grpTbls_StyleSet_TableQuote.Label = "Table &Quote Bullets Style Set"
        Me.grpTbls_StyleSet_TableQuote.Name = "grpTbls_StyleSet_TableQuote"
        Me.grpTbls_StyleSet_TableQuote.OfficeImageId = "FunctionsLogicalInsertGallery"
        Me.grpTbls_StyleSet_TableQuote.ScreenTip = "Table Quote Bullets Style Set"
        Me.grpTbls_StyleSet_TableQuote.ShowImage = True
        Me.grpTbls_StyleSet_TableQuote.SuperTip = """Make certain that your cursor is in a Table Cell """
        '
        'grpTbls_StyleSet_TableListBullets
        '
        Me.grpTbls_StyleSet_TableListBullets.Label = "Table List &Bullets Style Set"
        Me.grpTbls_StyleSet_TableListBullets.Name = "grpTbls_StyleSet_TableListBullets"
        Me.grpTbls_StyleSet_TableListBullets.OfficeImageId = "FunctionsLogicalInsertGallery"
        Me.grpTbls_StyleSet_TableListBullets.ScreenTip = "Table List Bullets Style Set"
        Me.grpTbls_StyleSet_TableListBullets.ShowImage = True
        Me.grpTbls_StyleSet_TableListBullets.SuperTip = """Make certain that your cursor is in a Table Cell """
        '
        'grpTbls_StyleSet_TableListNumbers
        '
        Me.grpTbls_StyleSet_TableListNumbers.Label = "Table List &Numbers Style Set"
        Me.grpTbls_StyleSet_TableListNumbers.Name = "grpTbls_StyleSet_TableListNumbers"
        Me.grpTbls_StyleSet_TableListNumbers.OfficeImageId = "FunctionsLogicalInsertGallery"
        Me.grpTbls_StyleSet_TableListNumbers.ScreenTip = "Table List Numbers Style Set"
        Me.grpTbls_StyleSet_TableListNumbers.ShowImage = True
        Me.grpTbls_StyleSet_TableListNumbers.SuperTip = """Make certain that your cursor is in a Table Cell """
        '
        'Separator50
        '
        Me.Separator50.Name = "Separator50"
        '
        'grpTbls_ColourCells
        '
        Me.grpTbls_ColourCells.Label = "Colour Cells"
        Me.grpTbls_ColourCells.Name = "grpTbls_ColourCells"
        Me.grpTbls_ColourCells.OfficeImageId = "FieldsMenu"
        Me.grpTbls_ColourCells.ScreenTip = "Colour Cells"
        Me.grpTbls_ColourCells.ShowImage = True
        Me.grpTbls_ColourCells.SuperTip = """Fills the selected cells with the same colour as the units row. Will work on tab" &
    "les with irregular structures"""
        Me.grpTbls_ColourCells.Visible = False
        '
        'grpTbls_ColourHeadingsRow
        '
        Me.grpTbls_ColourHeadingsRow.Label = "Colour Table Column Headings Row"
        Me.grpTbls_ColourHeadingsRow.Name = "grpTbls_ColourHeadingsRow"
        Me.grpTbls_ColourHeadingsRow.OfficeImageId = "RecordsMoreRecordsMenu"
        Me.grpTbls_ColourHeadingsRow.ScreenTip = "Colour Table Column Headings Row"
        Me.grpTbls_ColourHeadingsRow.ShowImage = True
        Me.grpTbls_ColourHeadingsRow.SuperTip = """Fills row with same colour as table heading row"""
        '
        'grpTbls_ColourUnitsRow
        '
        Me.grpTbls_ColourUnitsRow.Label = "Colour Table Units Row"
        Me.grpTbls_ColourUnitsRow.Name = "grpTbls_ColourUnitsRow"
        Me.grpTbls_ColourUnitsRow.OfficeImageId = "RecordsMoreRecordsMenu"
        Me.grpTbls_ColourUnitsRow.ScreenTip = "Colour Table Units Row"
        Me.grpTbls_ColourUnitsRow.ShowImage = True
        Me.grpTbls_ColourUnitsRow.SuperTip = """Fills row with same colour as table units row"""
        '
        'Separator51
        '
        Me.Separator51.Name = "Separator51"
        '
        'grpTbls_AllBorders
        '
        Me.grpTbls_AllBorders.Label = "Apply Table Borders to Row"
        Me.grpTbls_AllBorders.Name = "grpTbls_AllBorders"
        Me.grpTbls_AllBorders.OfficeImageId = "TableGrid2"
        Me.grpTbls_AllBorders.ScreenTip = "Apply Table Borders to Row"
        Me.grpTbls_AllBorders.ShowImage = True
        Me.grpTbls_AllBorders.SuperTip = """Applies all borders to the selected table. To select a table just place your cur" &
    "sor in it"""
        '
        'grpTbls_AllBordersRemove
        '
        Me.grpTbls_AllBordersRemove.Label = "Remove All Table Borders"
        Me.grpTbls_AllBordersRemove.Name = "grpTbls_AllBordersRemove"
        Me.grpTbls_AllBordersRemove.OfficeImageId = "TableGrid2"
        Me.grpTbls_AllBordersRemove.ScreenTip = "Remove All Table Borders"
        Me.grpTbls_AllBordersRemove.ShowImage = True
        Me.grpTbls_AllBordersRemove.SuperTip = """Removes all borders from the selected table. To select a table just place your c" &
    "ursor in it"""
        '
        'grpTbls_TableTextStyle
        '
        Me.grpTbls_TableTextStyle.KeyTip = "TTN"
        Me.grpTbls_TableTextStyle.Label = "Table text style"
        Me.grpTbls_TableTextStyle.Name = "grpTbls_TableTextStyle"
        Me.grpTbls_TableTextStyle.OfficeImageId = "BevelShapeGallery"
        Me.grpTbls_TableTextStyle.ScreenTip = "Table text style"
        Me.grpTbls_TableTextStyle.ShowImage = True
        Me.grpTbls_TableTextStyle.SuperTip = """Applies Table text style to selection"""
        '
        'mnuCloseDocuments4
        '
        Me.mnuCloseDocuments4.Items.Add(Me.grpTbls_convertTabletoES)
        Me.mnuCloseDocuments4.Items.Add(Me.grpTbls_convertTabletoStd)
        Me.mnuCloseDocuments4.Items.Add(Me.grpTbls_convertTabletoApp)
        Me.mnuCloseDocuments4.Items.Add(Me.Separator52)
        Me.mnuCloseDocuments4.Items.Add(Me.grpTbls_convertTabletoLT)
        Me.mnuCloseDocuments4.KeyTip = "T#"
        Me.mnuCloseDocuments4.Label = "Convert Table #"
        Me.mnuCloseDocuments4.Name = "mnuCloseDocuments4"
        Me.mnuCloseDocuments4.OfficeImageId = "DatasheetView"
        Me.mnuCloseDocuments4.ScreenTip = "Convert Table #"
        Me.mnuCloseDocuments4.ShowImage = True
        Me.mnuCloseDocuments4.SuperTip = """Use items from this menu to convert Table numbering captions to a different numb" &
    "ering format. You must select the entire caption for this function to work"""
        '
        'grpTbls_convertTabletoES
        '
        Me.grpTbls_convertTabletoES.Label = "Convert table to &ES numbering"
        Me.grpTbls_convertTabletoES.Name = "grpTbls_convertTabletoES"
        Me.grpTbls_convertTabletoES.OfficeImageId = "BevelShapeGallery"
        Me.grpTbls_convertTabletoES.ScreenTip = "Convert table to ES numbering"
        Me.grpTbls_convertTabletoES.ShowImage = True
        Me.grpTbls_convertTabletoES.SuperTip = """This function works on the individual placeholder caption. Select the entire Tab" &
    "le caption, then run this function to convert  to an ES numbered table."""
        '
        'grpTbls_convertTabletoStd
        '
        Me.grpTbls_convertTabletoStd.Label = "Convert Table to &Report numbering"
        Me.grpTbls_convertTabletoStd.Name = "grpTbls_convertTabletoStd"
        Me.grpTbls_convertTabletoStd.OfficeImageId = "BevelShapeGallery"
        Me.grpTbls_convertTabletoStd.ScreenTip = "Convert Table to Report numbering"
        Me.grpTbls_convertTabletoStd.ShowImage = True
        Me.grpTbls_convertTabletoStd.SuperTip = """This function works on individual placeholder caption. Select the entire Table c" &
    "aption, then run this function to convert to Report numbered table."""
        '
        'grpTbls_convertTabletoApp
        '
        Me.grpTbls_convertTabletoApp.Label = "Convert Table to &Appendix numbering"
        Me.grpTbls_convertTabletoApp.Name = "grpTbls_convertTabletoApp"
        Me.grpTbls_convertTabletoApp.OfficeImageId = "BevelShapeGallery"
        Me.grpTbls_convertTabletoApp.ScreenTip = "Convert Table to Appendix numbering"
        Me.grpTbls_convertTabletoApp.ShowImage = True
        Me.grpTbls_convertTabletoApp.SuperTip = """This function works on individual placeholder captions. Select the entire Table " &
    "caption, then click this button to convert to Appendix numbering."""
        '
        'Separator52
        '
        Me.Separator52.Name = "Separator52"
        '
        'grpTbls_convertTabletoLT
        '
        Me.grpTbls_convertTabletoLT.Label = "Convert Table to &Letter numbering"
        Me.grpTbls_convertTabletoLT.Name = "grpTbls_convertTabletoLT"
        Me.grpTbls_convertTabletoLT.OfficeImageId = "BevelShapeGallery"
        Me.grpTbls_convertTabletoLT.ScreenTip = "Convert Table to Letter numbering"
        Me.grpTbls_convertTabletoLT.ShowImage = True
        Me.grpTbls_convertTabletoLT.SuperTip = """This function works on individual placeholder captions. Select the entire Table " &
    "caption, then click this button to convert to Letter numbering."""
        '
        'grpTbls_AllStyles_small
        '
        Me.grpTbls_AllStyles_small.Items.Add(Me.grpTbls_TableColumnHeadingsStyle_small)
        Me.grpTbls_AllStyles_small.Items.Add(Me.grpTbls_TableUnitsRowStyle_small)
        Me.grpTbls_AllStyles_small.Items.Add(Me.Separator57)
        Me.grpTbls_AllStyles_small.Items.Add(Me.grpTbls_Plh_mnu_TableListBulletsStyles_small)
        Me.grpTbls_AllStyles_small.Items.Add(Me.grpTbls_Plh_mnu_TableListNumberingStyles_small)
        Me.grpTbls_AllStyles_small.Items.Add(Me.Separator58)
        Me.grpTbls_AllStyles_small.Items.Add(Me.grpPlh_mnu_TblSideHeadings_small)
        Me.grpTbls_AllStyles_small.Items.Add(Me.grpPlh_mnu_TblQuoteStyles_small)
        Me.grpTbls_AllStyles_small.Items.Add(Me.Separator59)
        Me.grpTbls_AllStyles_small.Items.Add(Me.grpTbl_Styles_ExampleStyleSets_Small)
        Me.grpTbls_AllStyles_small.KeyTip = "TM"
        Me.grpTbls_AllStyles_small.Label = "Table Styles (small)"
        Me.grpTbls_AllStyles_small.Name = "grpTbls_AllStyles_small"
        Me.grpTbls_AllStyles_small.OfficeImageId = "BevelShapeGallery"
        Me.grpTbls_AllStyles_small.ScreenTip = "Table Styles (small)"
        Me.grpTbls_AllStyles_small.ShowImage = True
        Me.grpTbls_AllStyles_small.SuperTip = """This menu provides access to a set of small font Table Styles.. These are typica" &
    "lly used when you wnat to put more data into a table of a given size"""
        '
        'grpTbls_TableColumnHeadingsStyle_small
        '
        Me.grpTbls_TableColumnHeadingsStyle_small.Label = "Table Column &Headings style"
        Me.grpTbls_TableColumnHeadingsStyle_small.Name = "grpTbls_TableColumnHeadingsStyle_small"
        Me.grpTbls_TableColumnHeadingsStyle_small.OfficeImageId = "AccessTableEvents"
        Me.grpTbls_TableColumnHeadingsStyle_small.ScreenTip = "Table Column Headings style"
        Me.grpTbls_TableColumnHeadingsStyle_small.ShowImage = True
        Me.grpTbls_TableColumnHeadingsStyle_small.SuperTip = "Applies Table Column Headings style to selection"
        '
        'grpTbls_TableUnitsRowStyle_small
        '
        Me.grpTbls_TableUnitsRowStyle_small.Label = "Table &Units Row style"
        Me.grpTbls_TableUnitsRowStyle_small.Name = "grpTbls_TableUnitsRowStyle_small"
        Me.grpTbls_TableUnitsRowStyle_small.OfficeImageId = "RecordsMoreRecordsMenu"
        Me.grpTbls_TableUnitsRowStyle_small.ScreenTip = "Table Units Row style"
        Me.grpTbls_TableUnitsRowStyle_small.ShowImage = True
        Me.grpTbls_TableUnitsRowStyle_small.SuperTip = "Applies Table Units Row style to selection"
        '
        'Separator57
        '
        Me.Separator57.Name = "Separator57"
        '
        'grpTbls_Plh_mnu_TableListBulletsStyles_small
        '
        Me.grpTbls_Plh_mnu_TableListBulletsStyles_small.Items.Add(Me.grpTbls_TableListBullet_small)
        Me.grpTbls_Plh_mnu_TableListBulletsStyles_small.Items.Add(Me.grpTbls_TableListBullet2_small)
        Me.grpTbls_Plh_mnu_TableListBulletsStyles_small.Items.Add(Me.grpTbls_TableListBullet3_small)
        Me.grpTbls_Plh_mnu_TableListBulletsStyles_small.Label = "Table List &Bullet Styles (small)"
        Me.grpTbls_Plh_mnu_TableListBulletsStyles_small.Name = "grpTbls_Plh_mnu_TableListBulletsStyles_small"
        Me.grpTbls_Plh_mnu_TableListBulletsStyles_small.OfficeImageId = "Bullets"
        Me.grpTbls_Plh_mnu_TableListBulletsStyles_small.ShowImage = True
        '
        'grpTbls_TableListBullet_small
        '
        Me.grpTbls_TableListBullet_small.Label = "Table List Bullet &1 style"
        Me.grpTbls_TableListBullet_small.Name = "grpTbls_TableListBullet_small"
        Me.grpTbls_TableListBullet_small.OfficeImageId = "Bullets"
        Me.grpTbls_TableListBullet_small.ShowImage = True
        Me.grpTbls_TableListBullet_small.SuperTip = "Small List Bullet"
        Me.grpTbls_TableListBullet_small.Tag = """Applies Table List Bullet style to selection"""
        '
        'grpTbls_TableListBullet2_small
        '
        Me.grpTbls_TableListBullet2_small.Label = "Table List Bullet &2 style"
        Me.grpTbls_TableListBullet2_small.Name = "grpTbls_TableListBullet2_small"
        Me.grpTbls_TableListBullet2_small.OfficeImageId = "Bullets"
        Me.grpTbls_TableListBullet2_small.ShowImage = True
        Me.grpTbls_TableListBullet2_small.SuperTip = "Small List Bullet"
        Me.grpTbls_TableListBullet2_small.Tag = """Applies Table List Bullet 2 (small) style to selection"""
        '
        'grpTbls_TableListBullet3_small
        '
        Me.grpTbls_TableListBullet3_small.Label = "Table List Bullet &3 style"
        Me.grpTbls_TableListBullet3_small.Name = "grpTbls_TableListBullet3_small"
        Me.grpTbls_TableListBullet3_small.OfficeImageId = "Bullets"
        Me.grpTbls_TableListBullet3_small.ScreenTip = "Small List Bullet"
        Me.grpTbls_TableListBullet3_small.ShowImage = True
        Me.grpTbls_TableListBullet3_small.SuperTip = """Applies Table List Bullet 3 (small) style to selection"""
        '
        'grpTbls_Plh_mnu_TableListNumberingStyles_small
        '
        Me.grpTbls_Plh_mnu_TableListNumberingStyles_small.Items.Add(Me.grpTbls_ListNumber_small)
        Me.grpTbls_Plh_mnu_TableListNumberingStyles_small.Items.Add(Me.grpTbls_ListNumber2_small)
        Me.grpTbls_Plh_mnu_TableListNumberingStyles_small.Items.Add(Me.grpTbls_ListNumber3_small)
        Me.grpTbls_Plh_mnu_TableListNumberingStyles_small.Label = "Table List &Number Styles (small)"
        Me.grpTbls_Plh_mnu_TableListNumberingStyles_small.Name = "grpTbls_Plh_mnu_TableListNumberingStyles_small"
        Me.grpTbls_Plh_mnu_TableListNumberingStyles_small.OfficeImageId = "Numbering"
        Me.grpTbls_Plh_mnu_TableListNumberingStyles_small.ShowImage = True
        '
        'grpTbls_ListNumber_small
        '
        Me.grpTbls_ListNumber_small.Label = "Table List Number &1 style"
        Me.grpTbls_ListNumber_small.Name = "grpTbls_ListNumber_small"
        Me.grpTbls_ListNumber_small.OfficeImageId = "Numbering"
        Me.grpTbls_ListNumber_small.ScreenTip = "Small List Number"
        Me.grpTbls_ListNumber_small.ShowImage = True
        Me.grpTbls_ListNumber_small.SuperTip = """Applies Table List Number (small) style to selection"""
        '
        'grpTbls_ListNumber2_small
        '
        Me.grpTbls_ListNumber2_small.Label = "Table List Number &2 style"
        Me.grpTbls_ListNumber2_small.Name = "grpTbls_ListNumber2_small"
        Me.grpTbls_ListNumber2_small.OfficeImageId = "Numbering"
        Me.grpTbls_ListNumber2_small.ScreenTip = "Small List Number"
        Me.grpTbls_ListNumber2_small.ShowImage = True
        Me.grpTbls_ListNumber2_small.SuperTip = """Applies Table List Number 2 (small) style to selection"""
        '
        'grpTbls_ListNumber3_small
        '
        Me.grpTbls_ListNumber3_small.Label = "Table List Number &3 style"
        Me.grpTbls_ListNumber3_small.Name = "grpTbls_ListNumber3_small"
        Me.grpTbls_ListNumber3_small.OfficeImageId = "Numbering"
        Me.grpTbls_ListNumber3_small.ScreenTip = "Small List Number"
        Me.grpTbls_ListNumber3_small.ShowImage = True
        Me.grpTbls_ListNumber3_small.SuperTip = """Applies Table List Number 3 (small) style to selection"""
        '
        'Separator58
        '
        Me.Separator58.Name = "Separator58"
        '
        'grpPlh_mnu_TblSideHeadings_small
        '
        Me.grpPlh_mnu_TblSideHeadings_small.Items.Add(Me.grpTbls_TableSideHeading1_small)
        Me.grpPlh_mnu_TblSideHeadings_small.Items.Add(Me.grpTbls_TableSideHeading2_small)
        Me.grpPlh_mnu_TblSideHeadings_small.Label = "Table &Side Heading (small)"
        Me.grpPlh_mnu_TblSideHeadings_small.Name = "grpPlh_mnu_TblSideHeadings_small"
        Me.grpPlh_mnu_TblSideHeadings_small.OfficeImageId = "PivotExpandField"
        Me.grpPlh_mnu_TblSideHeadings_small.ShowImage = True
        '
        'grpTbls_TableSideHeading1_small
        '
        Me.grpTbls_TableSideHeading1_small.Label = "Table Side Heading &1 (small)"
        Me.grpTbls_TableSideHeading1_small.Name = "grpTbls_TableSideHeading1_small"
        Me.grpTbls_TableSideHeading1_small.OfficeImageId = "PivotExpandField"
        Me.grpTbls_TableSideHeading1_small.ScreenTip = "Small Side Heading"
        Me.grpTbls_TableSideHeading1_small.ShowImage = True
        Me.grpTbls_TableSideHeading1_small.SuperTip = """Applies Table Side Heading 1 (small) style to selection"""
        '
        'grpTbls_TableSideHeading2_small
        '
        Me.grpTbls_TableSideHeading2_small.Label = "Table Side Heading &2 (small)"
        Me.grpTbls_TableSideHeading2_small.Name = "grpTbls_TableSideHeading2_small"
        Me.grpTbls_TableSideHeading2_small.OfficeImageId = "PivotExpandField"
        Me.grpTbls_TableSideHeading2_small.ScreenTip = "Small Side Heading"
        Me.grpTbls_TableSideHeading2_small.ShowImage = True
        Me.grpTbls_TableSideHeading2_small.SuperTip = """Applies Table Side Heading 2 (small) style to selection"""
        '
        'grpPlh_mnu_TblQuoteStyles_small
        '
        Me.grpPlh_mnu_TblQuoteStyles_small.Items.Add(Me.grpTbls_Quote_small)
        Me.grpPlh_mnu_TblQuoteStyles_small.Items.Add(Me.grpTbls_QuoteListBullet_small)
        Me.grpPlh_mnu_TblQuoteStyles_small.Items.Add(Me.grpTbls_QuoteSource_small)
        Me.grpPlh_mnu_TblQuoteStyles_small.Label = "Table &Quote Styles (small)"
        Me.grpPlh_mnu_TblQuoteStyles_small.Name = "grpPlh_mnu_TblQuoteStyles_small"
        Me.grpPlh_mnu_TblQuoteStyles_small.OfficeImageId = "BevelShapeGallery"
        Me.grpPlh_mnu_TblQuoteStyles_small.ScreenTip = "Table Quote Styles (small)"
        Me.grpPlh_mnu_TblQuoteStyles_small.ShowImage = True
        '
        'grpTbls_Quote_small
        '
        Me.grpTbls_Quote_small.Label = "Table &Quote style (small)"
        Me.grpTbls_Quote_small.Name = "grpTbls_Quote_small"
        Me.grpTbls_Quote_small.OfficeImageId = "BevelShapeGallery"
        Me.grpTbls_Quote_small.ScreenTip = "Table Quote style (small)"
        Me.grpTbls_Quote_small.ShowImage = True
        Me.grpTbls_Quote_small.SuperTip = """Applies Table Quote (small) style to selection"""
        '
        'grpTbls_QuoteListBullet_small
        '
        Me.grpTbls_QuoteListBullet_small.Label = "Table Quote List &Bullet (small) style"
        Me.grpTbls_QuoteListBullet_small.Name = "grpTbls_QuoteListBullet_small"
        Me.grpTbls_QuoteListBullet_small.OfficeImageId = "BevelShapeGallery"
        Me.grpTbls_QuoteListBullet_small.ScreenTip = "Table Quote List Bullet (small) style"
        Me.grpTbls_QuoteListBullet_small.ShowImage = True
        Me.grpTbls_QuoteListBullet_small.SuperTip = "Applies Table Quote List Bullet (small) style to selection"
        '
        'grpTbls_QuoteSource_small
        '
        Me.grpTbls_QuoteSource_small.Label = "Table &Quote Source (small) style"
        Me.grpTbls_QuoteSource_small.Name = "grpTbls_QuoteSource_small"
        Me.grpTbls_QuoteSource_small.OfficeImageId = "BevelShapeGallery"
        Me.grpTbls_QuoteSource_small.ScreenTip = "Table Quote Source (small) style"
        Me.grpTbls_QuoteSource_small.ShowImage = True
        Me.grpTbls_QuoteSource_small.SuperTip = """Applies Table Quote Source (small) style to selection"""
        '
        'Separator59
        '
        Me.Separator59.Name = "Separator59"
        '
        'grpTbl_Styles_ExampleStyleSets_Small
        '
        Me.grpTbl_Styles_ExampleStyleSets_Small.Items.Add(Me.grpTbls_StyleSet_TableQuote_small)
        Me.grpTbl_Styles_ExampleStyleSets_Small.Items.Add(Me.grpTbls_StyleSet_TableListBullets_small)
        Me.grpTbl_Styles_ExampleStyleSets_Small.Items.Add(Me.grpTbls_StyleSet_TableListNumbers_small)
        Me.grpTbl_Styles_ExampleStyleSets_Small.Label = "E&xample Style Sets (small)"
        Me.grpTbl_Styles_ExampleStyleSets_Small.Name = "grpTbl_Styles_ExampleStyleSets_Small"
        Me.grpTbl_Styles_ExampleStyleSets_Small.OfficeImageId = "FunctionsLogicalInsertGallery"
        Me.grpTbl_Styles_ExampleStyleSets_Small.ScreenTip = "Example Style Sets"
        Me.grpTbl_Styles_ExampleStyleSets_Small.ShowImage = True
        Me.grpTbl_Styles_ExampleStyleSets_Small.SuperTip = """This menu will allow you to insert example (small) style sets at the cursor posi" &
    "tion."""
        '
        'grpTbls_StyleSet_TableQuote_small
        '
        Me.grpTbls_StyleSet_TableQuote_small.Label = "Table Quote &Bullets (small) Style Set"
        Me.grpTbls_StyleSet_TableQuote_small.Name = "grpTbls_StyleSet_TableQuote_small"
        Me.grpTbls_StyleSet_TableQuote_small.OfficeImageId = "FunctionsLogicalInsertGallery"
        Me.grpTbls_StyleSet_TableQuote_small.ScreenTip = "Table Quote Bullets (small) Style Set"
        Me.grpTbls_StyleSet_TableQuote_small.ShowImage = True
        Me.grpTbls_StyleSet_TableQuote_small.SuperTip = """Will insert the Style Set in the selected Table Cell. """
        '
        'grpTbls_StyleSet_TableListBullets_small
        '
        Me.grpTbls_StyleSet_TableListBullets_small.Label = "Table &List Bullets (small) Style Set"
        Me.grpTbls_StyleSet_TableListBullets_small.Name = "grpTbls_StyleSet_TableListBullets_small"
        Me.grpTbls_StyleSet_TableListBullets_small.OfficeImageId = "FunctionsLogicalInsertGallery"
        Me.grpTbls_StyleSet_TableListBullets_small.ScreenTip = "List Bullets Style Set (small)"
        Me.grpTbls_StyleSet_TableListBullets_small.ShowImage = True
        Me.grpTbls_StyleSet_TableListBullets_small.SuperTip = """Will insert the Style Set in the selected Table Cell. """
        '
        'grpTbls_StyleSet_TableListNumbers_small
        '
        Me.grpTbls_StyleSet_TableListNumbers_small.Label = "Table List &Numbers (small) Style Set"
        Me.grpTbls_StyleSet_TableListNumbers_small.Name = "grpTbls_StyleSet_TableListNumbers_small"
        Me.grpTbls_StyleSet_TableListNumbers_small.OfficeImageId = "FunctionsLogicalInsertGallery"
        Me.grpTbls_StyleSet_TableListNumbers_small.ScreenTip = "List Numbers Style Set (small)"
        Me.grpTbls_StyleSet_TableListNumbers_small.ShowImage = True
        Me.grpTbls_StyleSet_TableListNumbers_small.SuperTip = """Will insert the Style Set in the selected Table Cell. """
        '
        'grpTbls_TableTextStyle_small
        '
        Me.grpTbls_TableTextStyle_small.KeyTip = "TTS"
        Me.grpTbls_TableTextStyle_small.Label = "Table text style (small)"
        Me.grpTbls_TableTextStyle_small.Name = "grpTbls_TableTextStyle_small"
        Me.grpTbls_TableTextStyle_small.OfficeImageId = "BevelShapeGallery"
        Me.grpTbls_TableTextStyle_small.ScreenTip = "Table text style (small)"
        Me.grpTbls_TableTextStyle_small.ShowImage = True
        Me.grpTbls_TableTextStyle_small.SuperTip = """Applies Table text (small) style to selection"""
        '
        'grpPlh_mnu_TblPlaceholders
        '
        Me.grpPlh_mnu_TblPlaceholders.Items.Add(Me.grpTblsPlh_HeadingAndSource)
        Me.grpPlh_mnu_TblPlaceholders.Items.Add(Me.grpTblsPlh_HeadingAndSourceApp)
        Me.grpPlh_mnu_TblPlaceholders.Items.Add(Me.grpTblsPlh_HeadingAndSourceES)
        Me.grpPlh_mnu_TblPlaceholders.Items.Add(Me.Separator64)
        Me.grpPlh_mnu_TblPlaceholders.Items.Add(Me.grpTblsPlh_CaptionAndHeading)
        Me.grpPlh_mnu_TblPlaceholders.Items.Add(Me.grpTblsPlh_CaptionAndHeadingApp)
        Me.grpPlh_mnu_TblPlaceholders.Items.Add(Me.grpTblsPlh_CaptionAndHeadingES)
        Me.grpPlh_mnu_TblPlaceholders.Items.Add(Me.Separator65)
        Me.grpPlh_mnu_TblPlaceholders.Items.Add(Me.grpTblsPlh_AddTable_Simple)
        Me.grpPlh_mnu_TblPlaceholders.KeyTip = "TP"
        Me.grpPlh_mnu_TblPlaceholders.Label = "Table placeholders"
        Me.grpPlh_mnu_TblPlaceholders.Name = "grpPlh_mnu_TblPlaceholders"
        Me.grpPlh_mnu_TblPlaceholders.OfficeImageId = "WindowSplit"
        Me.grpPlh_mnu_TblPlaceholders.ScreenTip = "Table placeholders"
        Me.grpPlh_mnu_TblPlaceholders.ShowImage = True
        Me.grpPlh_mnu_TblPlaceholders.SuperTip = resources.GetString("grpPlh_mnu_TblPlaceholders.SuperTip")
        '
        'grpTblsPlh_HeadingAndSource
        '
        Me.grpTblsPlh_HeadingAndSource.Label = "&Table Heading and Source"
        Me.grpTblsPlh_HeadingAndSource.Name = "grpTblsPlh_HeadingAndSource"
        Me.grpTblsPlh_HeadingAndSource.OfficeImageId = "WindowSplit"
        Me.grpTblsPlh_HeadingAndSource.ScreenTip = "Table Heading and Source"
        Me.grpTblsPlh_HeadingAndSource.ShowImage = True
        Me.grpTblsPlh_HeadingAndSource.SuperTip = "Inserts a standard placeholder with Table caption.  Paste table, or picture of ta" &
    "ble, into the empty space. "
        '
        'grpTblsPlh_HeadingAndSourceApp
        '
        Me.grpTblsPlh_HeadingAndSourceApp.Label = "&Appendix Table Heading and Source"
        Me.grpTblsPlh_HeadingAndSourceApp.Name = "grpTblsPlh_HeadingAndSourceApp"
        Me.grpTblsPlh_HeadingAndSourceApp.OfficeImageId = "WindowSplit"
        Me.grpTblsPlh_HeadingAndSourceApp.ScreenTip = "Appendix Table"
        Me.grpTblsPlh_HeadingAndSourceApp.ShowImage = True
        Me.grpTblsPlh_HeadingAndSourceApp.SuperTip = "Inserts a standard placeholder with Table caption prefixed by Appendix number.  P" &
    "aste table, or picture of table, into the empty space."
        '
        'grpTblsPlh_HeadingAndSourceES
        '
        Me.grpTblsPlh_HeadingAndSourceES.Label = "&ES Table Heading and Source"
        Me.grpTblsPlh_HeadingAndSourceES.Name = "grpTblsPlh_HeadingAndSourceES"
        Me.grpTblsPlh_HeadingAndSourceES.OfficeImageId = "WindowSplit"
        Me.grpTblsPlh_HeadingAndSourceES.ScreenTip = "ES Table"
        Me.grpTblsPlh_HeadingAndSourceES.ShowImage = True
        Me.grpTblsPlh_HeadingAndSourceES.SuperTip = "Inserts a standard placeholder with Table number prefixed by ES.  Use when a diff" &
    "erently numbered table is required in the Executive Summary. Paste table, or pic" &
    "ture of table, into the empty space."
        '
        'Separator64
        '
        Me.Separator64.Name = "Separator64"
        '
        'grpTblsPlh_CaptionAndHeading
        '
        Me.grpTblsPlh_CaptionAndHeading.Label = "Insert Table Caption and Heading only"
        Me.grpTblsPlh_CaptionAndHeading.Name = "grpTblsPlh_CaptionAndHeading"
        Me.grpTblsPlh_CaptionAndHeading.OfficeImageId = "WindowSplit"
        Me.grpTblsPlh_CaptionAndHeading.ScreenTip = "Caption and Heading"
        Me.grpTblsPlh_CaptionAndHeading.ShowImage = True
        Me.grpTblsPlh_CaptionAndHeading.SuperTip = "Inserts Table heading with Caption and number only"
        '
        'grpTblsPlh_CaptionAndHeadingApp
        '
        Me.grpTblsPlh_CaptionAndHeadingApp.Label = "Insert Appendix Table Caption and Heading only"
        Me.grpTblsPlh_CaptionAndHeadingApp.Name = "grpTblsPlh_CaptionAndHeadingApp"
        Me.grpTblsPlh_CaptionAndHeadingApp.OfficeImageId = "WindowSplit"
        Me.grpTblsPlh_CaptionAndHeadingApp.ScreenTip = "Appendix Caption and Heading"
        Me.grpTblsPlh_CaptionAndHeadingApp.ShowImage = True
        Me.grpTblsPlh_CaptionAndHeadingApp.SuperTip = "Inserts Appendix numbered Table Caption and number only."
        '
        'grpTblsPlh_CaptionAndHeadingES
        '
        Me.grpTblsPlh_CaptionAndHeadingES.Label = "Insert ES Table Caption and Heading only"
        Me.grpTblsPlh_CaptionAndHeadingES.Name = "grpTblsPlh_CaptionAndHeadingES"
        Me.grpTblsPlh_CaptionAndHeadingES.OfficeImageId = "WindowSplit"
        Me.grpTblsPlh_CaptionAndHeadingES.ScreenTip = "ES Caption and Heading"
        Me.grpTblsPlh_CaptionAndHeadingES.ShowImage = True
        Me.grpTblsPlh_CaptionAndHeadingES.SuperTip = "Inserts ES prefixed Table Caption and number only."
        '
        'Separator65
        '
        Me.Separator65.Name = "Separator65"
        '
        'grpTblsPlh_AddTable_Simple
        '
        Me.grpTblsPlh_AddTable_Simple.Label = "Add Table (&Simple)"
        Me.grpTblsPlh_AddTable_Simple.Name = "grpTblsPlh_AddTable_Simple"
        Me.grpTblsPlh_AddTable_Simple.OfficeImageId = "WindowSplit"
        Me.grpTblsPlh_AddTable_Simple.ScreenTip = "Table"
        Me.grpTblsPlh_AddTable_Simple.ShowImage = True
        Me.grpTblsPlh_AddTable_Simple.SuperTip = "Inserts the default Table from the Tables Gallery with rows and styles formatted " &
    "with the Table Text style."
        '
        'grpPlh_mnu_SourceAndNote
        '
        Me.grpPlh_mnu_SourceAndNote.Items.Add(Me.grpTblsPlh_SourceLabelAndStyle)
        Me.grpPlh_mnu_SourceAndNote.Items.Add(Me.grpTblsPlh_NoteLabelAndStyle)
        Me.grpPlh_mnu_SourceAndNote.Items.Add(Me.Separator66)
        Me.grpPlh_mnu_SourceAndNote.Items.Add(Me.grpTblsPlh_SourceForOverType)
        Me.grpPlh_mnu_SourceAndNote.KeyTip = "TN"
        Me.grpPlh_mnu_SourceAndNote.Label = "Source and Note"
        Me.grpPlh_mnu_SourceAndNote.Name = "grpPlh_mnu_SourceAndNote"
        Me.grpPlh_mnu_SourceAndNote.OfficeImageId = "WordCountList"
        Me.grpPlh_mnu_SourceAndNote.ScreenTip = "Source and Note"
        Me.grpPlh_mnu_SourceAndNote.ShowImage = True
        Me.grpPlh_mnu_SourceAndNote.SuperTip = """This menu contains preformatted note text which you can insert then overwrite as" &
    "  necessary. Apply the note styles to selected text if necessary."""
        '
        'grpTblsPlh_SourceLabelAndStyle
        '
        Me.grpTblsPlh_SourceLabelAndStyle.Label = "Insert &Source label"
        Me.grpTblsPlh_SourceLabelAndStyle.Name = "grpTblsPlh_SourceLabelAndStyle"
        Me.grpTblsPlh_SourceLabelAndStyle.OfficeImageId = "ReviewEditComment"
        Me.grpTblsPlh_SourceLabelAndStyle.ScreenTip = "Source label"
        Me.grpTblsPlh_SourceLabelAndStyle.ShowImage = True
        Me.grpTblsPlh_SourceLabelAndStyle.SuperTip = "Adds the source label at the selection and applies the source style to the paragr" &
    "aph"
        '
        'grpTblsPlh_NoteLabelAndStyle
        '
        Me.grpTblsPlh_NoteLabelAndStyle.Label = "Insert &Note label"
        Me.grpTblsPlh_NoteLabelAndStyle.Name = "grpTblsPlh_NoteLabelAndStyle"
        Me.grpTblsPlh_NoteLabelAndStyle.OfficeImageId = "ReviewEditComment"
        Me.grpTblsPlh_NoteLabelAndStyle.ScreenTip = "Note label"
        Me.grpTblsPlh_NoteLabelAndStyle.ShowImage = True
        Me.grpTblsPlh_NoteLabelAndStyle.SuperTip = "Adds the note label at the selection and applies the note style to the paragraph"
        '
        'Separator66
        '
        Me.Separator66.Name = "Separator66"
        '
        'grpTblsPlh_SourceForOverType
        '
        Me.grpTblsPlh_SourceForOverType.Label = "Source and Note for &overtyping"
        Me.grpTblsPlh_SourceForOverType.Name = "grpTblsPlh_SourceForOverType"
        Me.grpTblsPlh_SourceForOverType.OfficeImageId = "AdvertisePublishAs"
        Me.grpTblsPlh_SourceForOverType.ScreenTip = "For oevertyping"
        Me.grpTblsPlh_SourceForOverType.ShowImage = True
        Me.grpTblsPlh_SourceForOverType.SuperTip = "Inserts text to be overwritten as appropriate"
        '
        'grpPlh_mnu_DeleteTable
        '
        Me.grpPlh_mnu_DeleteTable.Items.Add(Me.grpTblsPlh_DeleteTable_fast)
        Me.grpPlh_mnu_DeleteTable.Items.Add(Me.grpTblsPlh_DeleteTable)
        Me.grpPlh_mnu_DeleteTable.KeyTip = "TX"
        Me.grpPlh_mnu_DeleteTable.Label = "Delete Table"
        Me.grpPlh_mnu_DeleteTable.Name = "grpPlh_mnu_DeleteTable"
        Me.grpPlh_mnu_DeleteTable.OfficeImageId = "TableDelete"
        Me.grpPlh_mnu_DeleteTable.ScreenTip = "Delete Table"
        Me.grpPlh_mnu_DeleteTable.ShowImage = True
        Me.grpPlh_mnu_DeleteTable.SuperTip = """The menu items here allow you to delete the selected Table. You need only place " &
    "your cursor anywhere in a Table to select it"""
        '
        'grpTblsPlh_DeleteTable_fast
        '
        Me.grpTblsPlh_DeleteTable_fast.Label = "Delete Table (no update - &fast)"
        Me.grpTblsPlh_DeleteTable_fast.Name = "grpTblsPlh_DeleteTable_fast"
        Me.grpTblsPlh_DeleteTable_fast.OfficeImageId = "BevelShapeGallery"
        Me.grpTblsPlh_DeleteTable_fast.ScreenTip = "Delete (no update)"
        Me.grpTblsPlh_DeleteTable_fast.ShowImage = True
        Me.grpTblsPlh_DeleteTable_fast.SuperTip = "This is a fast delete, it will delete the currently selected Table, but will not " &
    "update the document's 'Table Of Figures'."
        '
        'grpTblsPlh_DeleteTable
        '
        Me.grpTblsPlh_DeleteTable.Label = "Delete Table (with &update - slower)"
        Me.grpTblsPlh_DeleteTable.Name = "grpTblsPlh_DeleteTable"
        Me.grpTblsPlh_DeleteTable.OfficeImageId = "BevelShapeGallery"
        Me.grpTblsPlh_DeleteTable.ScreenTip = "Delete (with update)"
        Me.grpTblsPlh_DeleteTable.ShowImage = True
        Me.grpTblsPlh_DeleteTable.SuperTip = "This will delete the currently selected Table and will also update the document's" &
    " 'Table Of Figures', making it slower than the first option"
        '
        'grp_special_AATableFormatting
        '
        Me.grp_special_AATableFormatting.Items.Add(Me.tbPlh_mnu_convertPlhToHalfPage)
        Me.grp_special_AATableFormatting.Items.Add(Me.Separator69)
        Me.grp_special_AATableFormatting.Items.Add(Me.tbPlh_mnu_rapidFormat)
        Me.grp_special_AATableFormatting.Items.Add(Me.Separator70)
        Me.grp_special_AATableFormatting.Items.Add(Me.grpAATbls_mnu_editColumns)
        Me.grp_special_AATableFormatting.Items.Add(Me.grpAATbls_mnu_editRows)
        Me.grp_special_AATableFormatting.Items.Add(Me.grpAATbls_mnu_AATableactions)
        Me.grp_special_AATableFormatting.Items.Add(Me.Separator71)
        Me.grp_special_AATableFormatting.Items.Add(Me.grp_Plh_TableColumns_mnu_more)
        Me.grp_special_AATableFormatting.Label = "Special AA Table Formatting"
        Me.grp_special_AATableFormatting.Name = "grp_special_AATableFormatting"
        '
        'tbPlh_mnu_convertPlhToHalfPage
        '
        Me.tbPlh_mnu_convertPlhToHalfPage.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.tbPlh_mnu_convertPlhToHalfPage.Items.Add(Me.grpTbl_mnu_AAPlh_To_HalfPlh_Left)
        Me.tbPlh_mnu_convertPlhToHalfPage.Items.Add(Me.grpTbl_mnu_AAPlh_To_HalfPlh_Right)
        Me.tbPlh_mnu_convertPlhToHalfPage.Items.Add(Me.Separator67)
        Me.tbPlh_mnu_convertPlhToHalfPage.Items.Add(Me.grpTbl_mnu_AAPlh_Reset_to_FullColumn)
        Me.tbPlh_mnu_convertPlhToHalfPage.KeyTip = "CP2"
        Me.tbPlh_mnu_convertPlhToHalfPage.Label = "Convert Plh to 1/2 Page"
        Me.tbPlh_mnu_convertPlhToHalfPage.Name = "tbPlh_mnu_convertPlhToHalfPage"
        Me.tbPlh_mnu_convertPlhToHalfPage.OfficeImageId = "BevelShapeGallery"
        Me.tbPlh_mnu_convertPlhToHalfPage.ShowImage = True
        Me.tbPlh_mnu_convertPlhToHalfPage.SuperTip = resources.GetString("tbPlh_mnu_convertPlhToHalfPage.SuperTip")
        '
        'grpTbl_mnu_AAPlh_To_HalfPlh_Left
        '
        Me.grpTbl_mnu_AAPlh_To_HalfPlh_Left.Label = "&Left margin aligned half placeholder"
        Me.grpTbl_mnu_AAPlh_To_HalfPlh_Left.Name = "grpTbl_mnu_AAPlh_To_HalfPlh_Left"
        Me.grpTbl_mnu_AAPlh_To_HalfPlh_Left.OfficeImageId = "BevelShapeGallery"
        Me.grpTbl_mnu_AAPlh_To_HalfPlh_Left.ScreenTip = "Left aligned"
        Me.grpTbl_mnu_AAPlh_To_HalfPlh_Left.ShowImage = True
        Me.grpTbl_mnu_AAPlh_To_HalfPlh_Left.SuperTip = resources.GetString("grpTbl_mnu_AAPlh_To_HalfPlh_Left.SuperTip")
        '
        'grpTbl_mnu_AAPlh_To_HalfPlh_Right
        '
        Me.grpTbl_mnu_AAPlh_To_HalfPlh_Right.Label = "&Right margin aligned half placeholder"
        Me.grpTbl_mnu_AAPlh_To_HalfPlh_Right.Name = "grpTbl_mnu_AAPlh_To_HalfPlh_Right"
        Me.grpTbl_mnu_AAPlh_To_HalfPlh_Right.OfficeImageId = "BevelShapeGallery"
        Me.grpTbl_mnu_AAPlh_To_HalfPlh_Right.ScreenTip = "Right aligned"
        Me.grpTbl_mnu_AAPlh_To_HalfPlh_Right.ShowImage = True
        Me.grpTbl_mnu_AAPlh_To_HalfPlh_Right.SuperTip = resources.GetString("grpTbl_mnu_AAPlh_To_HalfPlh_Right.SuperTip")
        '
        'Separator67
        '
        Me.Separator67.Name = "Separator67"
        '
        'grpTbl_mnu_AAPlh_Reset_to_FullColumn
        '
        Me.grpTbl_mnu_AAPlh_Reset_to_FullColumn.Label = "&Reset to full column width"
        Me.grpTbl_mnu_AAPlh_Reset_to_FullColumn.Name = "grpTbl_mnu_AAPlh_Reset_to_FullColumn"
        Me.grpTbl_mnu_AAPlh_Reset_to_FullColumn.OfficeImageId = "BevelShapeGallery"
        Me.grpTbl_mnu_AAPlh_Reset_to_FullColumn.ScreenTip = "Reset to column"
        Me.grpTbl_mnu_AAPlh_Reset_to_FullColumn.ShowImage = True
        Me.grpTbl_mnu_AAPlh_Reset_to_FullColumn.SuperTip = "Place your cursor in an existing formatted Placeholder. &#13;&#13;This function w" &
    "ill adjust the Placeholder to inline and its &#13;&#13;width will be adjusted to" &
    " the full width of it's parent column."
        '
        'Separator69
        '
        Me.Separator69.Name = "Separator69"
        '
        'tbPlh_mnu_rapidFormat
        '
        Me.tbPlh_mnu_rapidFormat.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.tbPlh_mnu_rapidFormat.Items.Add(Me.grpTblsPlh_rapidFormat)
        Me.tbPlh_mnu_rapidFormat.Items.Add(Me.grpTblsPlh_rapidFormat_Encapsulated)
        Me.tbPlh_mnu_rapidFormat.Items.Add(Me.Separator68)
        Me.tbPlh_mnu_rapidFormat.Items.Add(Me.grpBoxes_mnu_rapidFormat_StdTbl_Force)
        Me.tbPlh_mnu_rapidFormat.Items.Add(Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force)
        Me.tbPlh_mnu_rapidFormat.KeyTip = "RF"
        Me.tbPlh_mnu_rapidFormat.Label = "Rapid Table Format"
        Me.tbPlh_mnu_rapidFormat.Name = "tbPlh_mnu_rapidFormat"
        Me.tbPlh_mnu_rapidFormat.OfficeImageId = "BevelShapeGallery"
        Me.tbPlh_mnu_rapidFormat.ShowImage = True
        '
        'grpTblsPlh_rapidFormat
        '
        Me.grpTblsPlh_rapidFormat.Label = "Standard table for &General use"
        Me.grpTblsPlh_rapidFormat.Name = "grpTblsPlh_rapidFormat"
        Me.grpTblsPlh_rapidFormat.OfficeImageId = "BevelShapeGallery"
        Me.grpTblsPlh_rapidFormat.ScreenTip = "Standard table"
        Me.grpTblsPlh_rapidFormat.ShowImage = True
        Me.grpTblsPlh_rapidFormat.SuperTip = resources.GetString("grpTblsPlh_rapidFormat.SuperTip")
        '
        'grpTblsPlh_rapidFormat_Encapsulated
        '
        Me.grpTblsPlh_rapidFormat_Encapsulated.Label = "Encapsulated table for &floating tables"
        Me.grpTblsPlh_rapidFormat_Encapsulated.Name = "grpTblsPlh_rapidFormat_Encapsulated"
        Me.grpTblsPlh_rapidFormat_Encapsulated.OfficeImageId = "BevelShapeGallery"
        Me.grpTblsPlh_rapidFormat_Encapsulated.ShowImage = True
        Me.grpTblsPlh_rapidFormat_Encapsulated.SuperTip = resources.GetString("grpTblsPlh_rapidFormat_Encapsulated.SuperTip")
        '
        'Separator68
        '
        Me.Separator68.Name = "Separator68"
        '
        'grpBoxes_mnu_rapidFormat_StdTbl_Force
        '
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force.Items.Add(Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_LT)
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force.Items.Add(Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_ES)
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force.Items.Add(Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_Body)
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force.Items.Add(Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_AP)
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force.Label = "&Standard Table (Selectable Caption)..."
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force.Name = "grpBoxes_mnu_rapidFormat_StdTbl_Force"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force.ScreenTip = "Standard Table"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force.ShowImage = True
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force.SuperTip = "Forces the auto generated table caption to the author's choice (LT, ES, Report Bo" &
    "dy or AP"
        '
        'grpBoxes_mnu_rapidFormat_StdTbl_Force_LT
        '
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_LT.Label = "Caption to '&LT' standard"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_LT.Name = "grpBoxes_mnu_rapidFormat_StdTbl_Force_LT"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_LT.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_LT.ScreenTip = "Caption to LT"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_LT.ShowImage = True
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_LT.SuperTip = "Place your cursor in an existing table. &#13;&#13;This function will adjust the T" &
    "able to the current AA format for general use and add the specified Caption at t" &
    "he top and Source at the bottom."
        '
        'grpBoxes_mnu_rapidFormat_StdTbl_Force_ES
        '
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_ES.Label = "Caption to '&ES' standard"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_ES.Name = "grpBoxes_mnu_rapidFormat_StdTbl_Force_ES"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_ES.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_ES.ScreenTip = "Caption to ES"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_ES.ShowImage = True
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_ES.SuperTip = "Place your cursor in an existing table. &#13;&#13;This function will adjust the T" &
    "able to the current AA format for general use and add the specified Caption at t" &
    "he top and Source at the bottom."
        '
        'grpBoxes_mnu_rapidFormat_StdTbl_Force_Body
        '
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_Body.Label = "Caption to '&Body' standard"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_Body.Name = "grpBoxes_mnu_rapidFormat_StdTbl_Force_Body"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_Body.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_Body.ScreenTip = "Caption to Body"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_Body.ShowImage = True
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_Body.SuperTip = "Place your cursor in an existing table. &#13;&#13;This function will adjust the T" &
    "able to the current AA format for general use and add the specified Caption at t" &
    "he top and Source at the bottom."
        '
        'grpBoxes_mnu_rapidFormat_StdTbl_Force_AP
        '
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_AP.Label = "Caption to '&AP' standard"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_AP.Name = "grpBoxes_mnu_rapidFormat_StdTbl_Force_AP"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_AP.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_AP.ScreenTip = "Caption to AP"
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_AP.ShowImage = True
        Me.grpBoxes_mnu_rapidFormat_StdTbl_Force_AP.SuperTip = "Place your cursor in an existing table. &#13;&#13;This function will adjust the T" &
    "able to the current AA format for general use and add the specified Caption at t" &
    "he top and Source at the bottom."
        '
        'grpBoxes_mnu_rapidFormat_EncapTbl_Force
        '
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force.Items.Add(Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_LT)
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force.Items.Add(Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_ES)
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force.Items.Add(Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_Body)
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force.Items.Add(Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_AP)
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force.Label = "&Encapsulated Table (Selectable Caption)..."
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force.Name = "grpBoxes_mnu_rapidFormat_EncapTbl_Force"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force.ScreenTip = "Encapsulated Table"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force.ShowImage = True
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force.SuperTip = "Forces the auto generated table caption to the author's choice (LT, ES, Report Bo" &
    "dy or AP"
        '
        'grpBoxes_mnu_rapidFormat_EncapTbl_Force_LT
        '
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_LT.Label = "Caption to '&LT' standard"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_LT.Name = "grpBoxes_mnu_rapidFormat_EncapTbl_Force_LT"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_LT.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_LT.ScreenTip = "Caption to LT"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_LT.ShowImage = True
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_LT.SuperTip = resources.GetString("grpBoxes_mnu_rapidFormat_EncapTbl_Force_LT.SuperTip")
        '
        'grpBoxes_mnu_rapidFormat_EncapTbl_Force_ES
        '
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_ES.Label = "Caption to '&ES' standard"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_ES.Name = "grpBoxes_mnu_rapidFormat_EncapTbl_Force_ES"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_ES.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_ES.ScreenTip = "Caption to ES"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_ES.ShowImage = True
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_ES.SuperTip = resources.GetString("grpBoxes_mnu_rapidFormat_EncapTbl_Force_ES.SuperTip")
        '
        'grpBoxes_mnu_rapidFormat_EncapTbl_Force_Body
        '
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_Body.Label = "Caption to '&Body' standard"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_Body.Name = "grpBoxes_mnu_rapidFormat_EncapTbl_Force_Body"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_Body.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_Body.ScreenTip = "Caption to Body"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_Body.ShowImage = True
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_Body.SuperTip = resources.GetString("grpBoxes_mnu_rapidFormat_EncapTbl_Force_Body.SuperTip")
        '
        'grpBoxes_mnu_rapidFormat_EncapTbl_Force_AP
        '
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_AP.Label = "Caption to '&AP' standard"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_AP.Name = "grpBoxes_mnu_rapidFormat_EncapTbl_Force_AP"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_AP.OfficeImageId = "BevelShapeGallery"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_AP.ScreenTip = "Caption to AP"
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_AP.ShowImage = True
        Me.grpBoxes_mnu_rapidFormat_EncapTbl_Force_AP.SuperTip = resources.GetString("grpBoxes_mnu_rapidFormat_EncapTbl_Force_AP.SuperTip")
        '
        'Separator70
        '
        Me.Separator70.Name = "Separator70"
        '
        'grpAATbls_mnu_editColumns
        '
        Me.grpAATbls_mnu_editColumns.Items.Add(Me.grpTblsEdit_InsertColumnRight)
        Me.grpAATbls_mnu_editColumns.Items.Add(Me.grpTblsEdit_InsertColumnLeft)
        Me.grpAATbls_mnu_editColumns.Items.Add(Me.Separator72)
        Me.grpAATbls_mnu_editColumns.Items.Add(Me.grpTblsEdit_Delete_Column)
        Me.grpAATbls_mnu_editColumns.KeyTip = "ATC"
        Me.grpAATbls_mnu_editColumns.Label = "Edit AA Columns (Beta)"
        Me.grpAATbls_mnu_editColumns.Name = "grpAATbls_mnu_editColumns"
        Me.grpAATbls_mnu_editColumns.OfficeImageId = "TableColumnSelect"
        Me.grpAATbls_mnu_editColumns.ScreenTip = "Functions to insert and delete columns to/from AA standard and encapsulated table" &
    "s."
        Me.grpAATbls_mnu_editColumns.ShowImage = True
        Me.grpAATbls_mnu_editColumns.SuperTip = resources.GetString("grpAATbls_mnu_editColumns.SuperTip")
        '
        'grpTblsEdit_InsertColumnRight
        '
        Me.grpTblsEdit_InsertColumnRight.Label = "Insert &Right"
        Me.grpTblsEdit_InsertColumnRight.Name = "grpTblsEdit_InsertColumnRight"
        Me.grpTblsEdit_InsertColumnRight.OfficeImageId = "InsertColumnRight"
        Me.grpTblsEdit_InsertColumnRight.ScreenTip = "Right"
        Me.grpTblsEdit_InsertColumnRight.ShowImage = True
        Me.grpTblsEdit_InsertColumnRight.SuperTip = resources.GetString("grpTblsEdit_InsertColumnRight.SuperTip")
        '
        'grpTblsEdit_InsertColumnLeft
        '
        Me.grpTblsEdit_InsertColumnLeft.Label = "Insert &Left"
        Me.grpTblsEdit_InsertColumnLeft.Name = "grpTblsEdit_InsertColumnLeft"
        Me.grpTblsEdit_InsertColumnLeft.OfficeImageId = "InsertColumnLeft"
        Me.grpTblsEdit_InsertColumnLeft.ScreenTip = "Left"
        Me.grpTblsEdit_InsertColumnLeft.ShowImage = True
        Me.grpTblsEdit_InsertColumnLeft.SuperTip = resources.GetString("grpTblsEdit_InsertColumnLeft.SuperTip")
        '
        'Separator72
        '
        Me.Separator72.Name = "Separator72"
        '
        'grpTblsEdit_Delete_Column
        '
        Me.grpTblsEdit_Delete_Column.Label = "&Delete Column"
        Me.grpTblsEdit_Delete_Column.Name = "grpTblsEdit_Delete_Column"
        Me.grpTblsEdit_Delete_Column.OfficeImageId = "DeleteColumns"
        Me.grpTblsEdit_Delete_Column.ScreenTip = "Delete"
        Me.grpTblsEdit_Delete_Column.ShowImage = True
        Me.grpTblsEdit_Delete_Column.SuperTip = "Place your cursor in the table column of interest. &#13;&#13;This function will w" &
    "ork for both 'standard' and 'encapsulated' AA tables."
        '
        'grpAATbls_mnu_editRows
        '
        Me.grpAATbls_mnu_editRows.Items.Add(Me.grpTblsEdit_InsertRowAbove)
        Me.grpAATbls_mnu_editRows.Items.Add(Me.grpTblsEdit_InsertRowBelow)
        Me.grpAATbls_mnu_editRows.Items.Add(Me.Separator73)
        Me.grpAATbls_mnu_editRows.Items.Add(Me.grpTblsEdit_Delete_Row)
        Me.grpAATbls_mnu_editRows.KeyTip = "ATR"
        Me.grpAATbls_mnu_editRows.Label = "Edit AA Rows (Beta)"
        Me.grpAATbls_mnu_editRows.Name = "grpAATbls_mnu_editRows"
        Me.grpAATbls_mnu_editRows.OfficeImageId = "TableRowSelect"
        Me.grpAATbls_mnu_editRows.ScreenTip = "Functions to insert and delete rows."
        Me.grpAATbls_mnu_editRows.ShowImage = True
        Me.grpAATbls_mnu_editRows.SuperTip = "If something drastic happens when editing the rows, a function that will paste ba" &
    "ck the prior version of the table over the selected table."
        '
        'grpTblsEdit_InsertRowAbove
        '
        Me.grpTblsEdit_InsertRowAbove.Label = "Insert &Above"
        Me.grpTblsEdit_InsertRowAbove.Name = "grpTblsEdit_InsertRowAbove"
        Me.grpTblsEdit_InsertRowAbove.OfficeImageId = "InsertRowAbove"
        Me.grpTblsEdit_InsertRowAbove.ScreenTip = "Place your cursor in a row of the target AA Table (Standard or Encapsulated)"
        Me.grpTblsEdit_InsertRowAbove.ShowImage = True
        Me.grpTblsEdit_InsertRowAbove.SuperTip = "If something drastic happens when editing the rows, a function that will paste ba" &
    "ck the prior version of the table over the selected table."
        '
        'grpTblsEdit_InsertRowBelow
        '
        Me.grpTblsEdit_InsertRowBelow.Label = "Insert &Below"
        Me.grpTblsEdit_InsertRowBelow.Name = "grpTblsEdit_InsertRowBelow"
        Me.grpTblsEdit_InsertRowBelow.OfficeImageId = "InsertRowBelow"
        Me.grpTblsEdit_InsertRowBelow.ScreenTip = "Insert below"
        Me.grpTblsEdit_InsertRowBelow.ShowImage = True
        Me.grpTblsEdit_InsertRowBelow.SuperTip = "Place your cursor in a row of the target AA Table (Standard or Encapsulated). Thi" &
    "s function will insert a row below the row that contains your cursor."
        '
        'Separator73
        '
        Me.Separator73.Name = "Separator73"
        '
        'grpTblsEdit_Delete_Row
        '
        Me.grpTblsEdit_Delete_Row.Label = "&Delete Row"
        Me.grpTblsEdit_Delete_Row.Name = "grpTblsEdit_Delete_Row"
        Me.grpTblsEdit_Delete_Row.OfficeImageId = "DeleteRows"
        Me.grpTblsEdit_Delete_Row.ScreenTip = "Place your cursor in the table column of interest."
        Me.grpTblsEdit_Delete_Row.ShowImage = True
        Me.grpTblsEdit_Delete_Row.SuperTip = "This function will work for both 'standard' and 'encapsulated' AA tables, but the" &
    "y do need to be reular by column."
        '
        'grpAATbls_mnu_AATableactions
        '
        Me.grpAATbls_mnu_AATableactions.Items.Add(Me.grpTblsEdit_CopyTable)
        Me.grpAATbls_mnu_AATableactions.Items.Add(Me.Separator74)
        Me.grpAATbls_mnu_AATableactions.Items.Add(Me.grpTblsEdit_PastePriorTable)
        Me.grpAATbls_mnu_AATableactions.Items.Add(Me.Separator75)
        Me.grpAATbls_mnu_AATableactions.Items.Add(Me.grpTblsEdit_UndoTableAction)
        Me.grpAATbls_mnu_AATableactions.KeyTip = "ATA"
        Me.grpAATbls_mnu_AATableactions.Label = "AA Table Actions"
        Me.grpAATbls_mnu_AATableactions.Name = "grpAATbls_mnu_AATableactions"
        Me.grpAATbls_mnu_AATableactions.OfficeImageId = "TablesResetToDefault"
        Me.grpAATbls_mnu_AATableactions.ScreenTip = "Provides functions to insert and delete columns/rows from AA Tables."
        Me.grpAATbls_mnu_AATableactions.ShowImage = True
        Me.grpAATbls_mnu_AATableactions.SuperTip = "It also provides a function to paste a prior version of the current working table" &
    " back over the modified table (Undo), or into a free area of the document (Paste" &
    ")."
        '
        'grpTblsEdit_CopyTable
        '
        Me.grpTblsEdit_CopyTable.Label = "&Copy Table"
        Me.grpTblsEdit_CopyTable.Name = "grpTblsEdit_CopyTable"
        Me.grpTblsEdit_CopyTable.OfficeImageId = "TablesMore"
        Me.grpTblsEdit_CopyTable.ScreenTip = "Copy Table"
        Me.grpTblsEdit_CopyTable.ShowImage = True
        Me.grpTblsEdit_CopyTable.SuperTip = "This function will copy the current table (your cursor must be somewhere in the t" &
    "able) to the Clipboard."
        '
        'Separator74
        '
        Me.Separator74.Name = "Separator74"
        '
        'grpTblsEdit_PastePriorTable
        '
        Me.grpTblsEdit_PastePriorTable.Label = "&Paste AA Table"
        Me.grpTblsEdit_PastePriorTable.Name = "grpTblsEdit_PastePriorTable"
        Me.grpTblsEdit_PastePriorTable.OfficeImageId = "PasteSingleCellTableAsTable"
        Me.grpTblsEdit_PastePriorTable.ScreenTip = "Paste Table"
        Me.grpTblsEdit_PastePriorTable.ShowImage = True
        Me.grpTblsEdit_PastePriorTable.SuperTip = "If you use this immediately after using the 'AA Column/Row' edit functions it wil" &
    "l paste the 'prior' AA Table at the current cursor position."
        '
        'Separator75
        '
        Me.Separator75.Name = "Separator75"
        '
        'grpTblsEdit_UndoTableAction
        '
        Me.grpTblsEdit_UndoTableAction.Label = "&Undo AA Table edit"
        Me.grpTblsEdit_UndoTableAction.Name = "grpTblsEdit_UndoTableAction"
        Me.grpTblsEdit_UndoTableAction.OfficeImageId = "Undo"
        Me.grpTblsEdit_UndoTableAction.ScreenTip = "Undo"
        Me.grpTblsEdit_UndoTableAction.ShowImage = True
        Me.grpTblsEdit_UndoTableAction.SuperTip = resources.GetString("grpTblsEdit_UndoTableAction.SuperTip")
        '
        'Separator71
        '
        Me.Separator71.Name = "Separator71"
        '
        'grp_Plh_TableColumns_mnu_more
        '
        Me.grp_Plh_TableColumns_mnu_more.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grp_Plh_TableColumns_mnu_more.Items.Add(Me.grpTblsEdit_Convert_EncapsToStd)
        Me.grp_Plh_TableColumns_mnu_more.Items.Add(Me.grpTblsEdit_Convert_StdToEncaps)
        Me.grp_Plh_TableColumns_mnu_more.Items.Add(Me.Separator76)
        Me.grp_Plh_TableColumns_mnu_more.Items.Add(Me.grpTblsEdit_Split_Table)
        Me.grp_Plh_TableColumns_mnu_more.KeyTip = "TTC"
        Me.grp_Plh_TableColumns_mnu_more.Label = "Table type conversions"
        Me.grp_Plh_TableColumns_mnu_more.Name = "grp_Plh_TableColumns_mnu_more"
        Me.grp_Plh_TableColumns_mnu_more.OfficeImageId = "TabSendReceive"
        Me.grp_Plh_TableColumns_mnu_more.ScreenTip = "Tools that support conversion between AA standard and AA Encapsulated tables. "
        Me.grp_Plh_TableColumns_mnu_more.ShowImage = True
        Me.grp_Plh_TableColumns_mnu_more.SuperTip = "You can also split an Encapsulated Table, allowing you to manually add rows, colu" &
    "mns and merged cells... This is for those times when the 'Edit Columns/Rows' fun" &
    "ction to the left do not work."
        '
        'grpTblsEdit_Convert_EncapsToStd
        '
        Me.grpTblsEdit_Convert_EncapsToStd.Label = "Convert to &Standard Table"
        Me.grpTblsEdit_Convert_EncapsToStd.Name = "grpTblsEdit_Convert_EncapsToStd"
        Me.grpTblsEdit_Convert_EncapsToStd.OfficeImageId = "RuleLinesLargeGrid"
        Me.grpTblsEdit_Convert_EncapsToStd.ScreenTip = "Encapsulated to Standard"
        Me.grpTblsEdit_Convert_EncapsToStd.ShowImage = True
        Me.grpTblsEdit_Convert_EncapsToStd.SuperTip = "This function will convert a AA Encapsulated Table to an AA Standard Table... Pla" &
    "ce your cursor in the table of interest."
        '
        'grpTblsEdit_Convert_StdToEncaps
        '
        Me.grpTblsEdit_Convert_StdToEncaps.Label = "Convert to &Encapsulated Table"
        Me.grpTblsEdit_Convert_StdToEncaps.Name = "grpTblsEdit_Convert_StdToEncaps"
        Me.grpTblsEdit_Convert_StdToEncaps.OfficeImageId = "TableInsertDialog"
        Me.grpTblsEdit_Convert_StdToEncaps.ScreenTip = "Standard to Encapsulated"
        Me.grpTblsEdit_Convert_StdToEncaps.ShowImage = True
        Me.grpTblsEdit_Convert_StdToEncaps.SuperTip = "This function will convert an AA Standard Table to an AA Encapsulated Table... Pl" &
    "ace your cursor in the table of interest."
        '
        'Separator76
        '
        Me.Separator76.Name = "Separator76"
        '
        'grpTblsEdit_Split_Table
        '
        Me.grpTblsEdit_Split_Table.Label = "Sp&lit an Encapsulated Table"
        Me.grpTblsEdit_Split_Table.Name = "grpTblsEdit_Split_Table"
        Me.grpTblsEdit_Split_Table.OfficeImageId = "TableSplitTable"
        Me.grpTblsEdit_Split_Table.ScreenTip = "Split and Encapsulated Table"
        Me.grpTblsEdit_Split_Table.ShowImage = True
        Me.grpTblsEdit_Split_Table.SuperTip = "This function will split a AA Encapsulated Table into three parts. That is, the t" &
    "op caption row, the body and the bottom source row... Place your cursor in the t" &
    "able of interest."
        '
        'grp_floatingPlaceholders
        '
        Me.grp_floatingPlaceholders.Items.Add(Me.grpReport_PlH_Handling)
        Me.grp_floatingPlaceholders.Items.Add(Me.grpReport_PlH_convertToInline_findAllFloatingTables_2)
        Me.grp_floatingPlaceholders.Items.Add(Me.grpReport_PlH_convertToInline)
        Me.grp_floatingPlaceholders.Label = "Floating Placeholders"
        Me.grp_floatingPlaceholders.Name = "grp_floatingPlaceholders"
        '
        'grpReport_PlH_Handling
        '
        Me.grpReport_PlH_Handling.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grpReport_PlH_Handling.Items.Add(Me.grpReport_PlH_LockToTop)
        Me.grpReport_PlH_Handling.Items.Add(Me.grpReport_PlHFloat_lock_toMarginsLeftAndTop)
        Me.grpReport_PlH_Handling.Items.Add(Me.grpReport_PlH_LockToParagraph)
        Me.grpReport_PlH_Handling.Items.Add(Me.grpReport_PlH_LockToParagraphAndColumn)
        Me.grpReport_PlH_Handling.Items.Add(Me.Separator77)
        Me.grpReport_PlH_Handling.Items.Add(Me.grpReport_PlH_FloatEdgeToEdge)
        Me.grpReport_PlH_Handling.Items.Add(Me.grpReport_PlH_FloatWide)
        Me.grpReport_PlH_Handling.Items.Add(Me.grpReport_PlH_FloatMarginToMargin)
        Me.grpReport_PlH_Handling.Items.Add(Me.grpReport_PlH_ColumnWidth)
        Me.grpReport_PlH_Handling.Items.Add(Me.grpReport_PlH_TwoColumnWidth)
        Me.grpReport_PlH_Handling.KeyTip = "LM"
        Me.grpReport_PlH_Handling.Label = "Placeholder Mgmnt"
        Me.grpReport_PlH_Handling.Name = "grpReport_PlH_Handling"
        Me.grpReport_PlH_Handling.OfficeImageId = "AssetSettings"
        Me.grpReport_PlH_Handling.ScreenTip = "Place holder management tools for floating tables, boxes and figures"
        Me.grpReport_PlH_Handling.ShowImage = True
        Me.grpReport_PlH_Handling.SuperTip = resources.GetString("grpReport_PlH_Handling.SuperTip")
        '
        'grpReport_PlH_LockToTop
        '
        Me.grpReport_PlH_LockToTop.Label = "Float and Lock to &Top Margin"
        Me.grpReport_PlH_LockToTop.Name = "grpReport_PlH_LockToTop"
        Me.grpReport_PlH_LockToTop.OfficeImageId = "ManuallyScheduleSelectedTask"
        Me.grpReport_PlH_LockToTop.ScreenTip = "Float and Lock to Top Margin"
        Me.grpReport_PlH_LockToTop.ShowImage = True
        '
        'grpReport_PlHFloat_lock_toMarginsLeftAndTop
        '
        Me.grpReport_PlHFloat_lock_toMarginsLeftAndTop.Label = "Float and Lock relative to top and left margins"
        Me.grpReport_PlHFloat_lock_toMarginsLeftAndTop.Name = "grpReport_PlHFloat_lock_toMarginsLeftAndTop"
        Me.grpReport_PlHFloat_lock_toMarginsLeftAndTop.OfficeImageId = "ManuallyScheduleSelectedTask"
        Me.grpReport_PlHFloat_lock_toMarginsLeftAndTop.ScreenTip = "Lock relative to top and left margins"
        Me.grpReport_PlHFloat_lock_toMarginsLeftAndTop.ShowImage = True
        Me.grpReport_PlHFloat_lock_toMarginsLeftAndTop.SuperTip = resources.GetString("grpReport_PlHFloat_lock_toMarginsLeftAndTop.SuperTip")
        '
        'grpReport_PlH_LockToParagraph
        '
        Me.grpReport_PlH_LockToParagraph.Label = "Float and Lock to the left margin and a &paragraph"
        Me.grpReport_PlH_LockToParagraph.Name = "grpReport_PlH_LockToParagraph"
        Me.grpReport_PlH_LockToParagraph.OfficeImageId = "ManuallyScheduleSelectedTask"
        Me.grpReport_PlH_LockToParagraph.ScreenTip = "Lock to left margin and paragarph"
        Me.grpReport_PlH_LockToParagraph.ShowImage = True
        Me.grpReport_PlH_LockToParagraph.SuperTip = resources.GetString("grpReport_PlH_LockToParagraph.SuperTip")
        '
        'grpReport_PlH_LockToParagraphAndColumn
        '
        Me.grpReport_PlH_LockToParagraphAndColumn.Label = "Float and Lock relative to a column and a paragraph"
        Me.grpReport_PlH_LockToParagraphAndColumn.Name = "grpReport_PlH_LockToParagraphAndColumn"
        Me.grpReport_PlH_LockToParagraphAndColumn.OfficeImageId = "ManuallyScheduleSelectedTask"
        Me.grpReport_PlH_LockToParagraphAndColumn.ScreenTip = "Float and Lock relative to a column and a paragraph"
        Me.grpReport_PlH_LockToParagraphAndColumn.ShowImage = True
        Me.grpReport_PlH_LockToParagraphAndColumn.SuperTip = resources.GetString("grpReport_PlH_LockToParagraphAndColumn.SuperTip")
        '
        'Separator77
        '
        Me.Separator77.Name = "Separator77"
        '
        'grpReport_PlH_FloatEdgeToEdge
        '
        Me.grpReport_PlH_FloatEdgeToEdge.Label = "Fit between the page &edges"
        Me.grpReport_PlH_FloatEdgeToEdge.Name = "grpReport_PlH_FloatEdgeToEdge"
        Me.grpReport_PlH_FloatEdgeToEdge.OfficeImageId = "AsianLayoutFitText"
        Me.grpReport_PlH_FloatEdgeToEdge.ScreenTip = "Fit between the page edges"
        Me.grpReport_PlH_FloatEdgeToEdge.ShowImage = True
        Me.grpReport_PlH_FloatEdgeToEdge.SuperTip = resources.GetString("grpReport_PlH_FloatEdgeToEdge.SuperTip")
        '
        'grpReport_PlH_FloatWide
        '
        Me.grpReport_PlH_FloatWide.Label = "Fit between the header table edges"
        Me.grpReport_PlH_FloatWide.Name = "grpReport_PlH_FloatWide"
        Me.grpReport_PlH_FloatWide.OfficeImageId = "AsianLayoutFitText"
        Me.grpReport_PlH_FloatWide.ScreenTip = "Fit between the header table edges"
        Me.grpReport_PlH_FloatWide.ShowImage = True
        Me.grpReport_PlH_FloatWide.SuperTip = "Will adjust the selected Placeholder's width relative to the left and right edges" &
    " of the 'header' table... Make certain that your selection point is in the Place" &
    "holder of interest"
        '
        'grpReport_PlH_FloatMarginToMargin
        '
        Me.grpReport_PlH_FloatMarginToMargin.Label = "Fit between the &margins"
        Me.grpReport_PlH_FloatMarginToMargin.Name = "grpReport_PlH_FloatMarginToMargin"
        Me.grpReport_PlH_FloatMarginToMargin.OfficeImageId = "AsianLayoutFitText"
        Me.grpReport_PlH_FloatMarginToMargin.ScreenTip = "Fit between the margins"
        Me.grpReport_PlH_FloatMarginToMargin.ShowImage = True
        Me.grpReport_PlH_FloatMarginToMargin.SuperTip = "Will adjust the selected Placeholder's width relative to the left and right margi" &
    "ns of the page... Make certain that your selection point is in the Placeholder o" &
    "f interest"
        '
        'grpReport_PlH_ColumnWidth
        '
        Me.grpReport_PlH_ColumnWidth.Label = "Fit to the &parent column"
        Me.grpReport_PlH_ColumnWidth.Name = "grpReport_PlH_ColumnWidth"
        Me.grpReport_PlH_ColumnWidth.OfficeImageId = "AsianLayoutFitText"
        Me.grpReport_PlH_ColumnWidth.ScreenTip = "Fit to the parent column"
        Me.grpReport_PlH_ColumnWidth.ShowImage = True
        Me.grpReport_PlH_ColumnWidth.SuperTip = "Will adjust the selected Placeholder's width relative to the left and right edges" &
    " of the parent column... Make certain that your selection point is in the Placeh" &
    "older of interest."
        '
        'grpReport_PlH_TwoColumnWidth
        '
        Me.grpReport_PlH_TwoColumnWidth.Label = "Fit to &2 columns"
        Me.grpReport_PlH_TwoColumnWidth.Name = "grpReport_PlH_TwoColumnWidth"
        Me.grpReport_PlH_TwoColumnWidth.OfficeImageId = "AsianLayoutFitText"
        Me.grpReport_PlH_TwoColumnWidth.ScreenTip = "Expand across two columns"
        Me.grpReport_PlH_TwoColumnWidth.ShowImage = True
        Me.grpReport_PlH_TwoColumnWidth.SuperTip = resources.GetString("grpReport_PlH_TwoColumnWidth.SuperTip")
        '
        'grpReport_PlH_convertToInline_findAllFloatingTables_2
        '
        Me.grpReport_PlH_convertToInline_findAllFloatingTables_2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grpReport_PlH_convertToInline_findAllFloatingTables_2.KeyTip = "PM"
        Me.grpReport_PlH_convertToInline_findAllFloatingTables_2.Label = "Placeholder Map"
        Me.grpReport_PlH_convertToInline_findAllFloatingTables_2.Name = "grpReport_PlH_convertToInline_findAllFloatingTables_2"
        Me.grpReport_PlH_convertToInline_findAllFloatingTables_2.OfficeImageId = "MapContactAddress"
        Me.grpReport_PlH_convertToInline_findAllFloatingTables_2.ScreenTip = "'Accessible' documents do NOT permit 'Floating' or 'Irregular' Tables/Placeholder" &
    "s."
        Me.grpReport_PlH_convertToInline_findAllFloatingTables_2.ShowImage = True
        Me.grpReport_PlH_convertToInline_findAllFloatingTables_2.SuperTip = "This form provides tools that allow you to go to, or convert a placeholder from f" &
    "loating to inline. Or to find and easily go to Irregular tables."
        '
        'grpReport_PlH_convertToInline
        '
        Me.grpReport_PlH_convertToInline.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grpReport_PlH_convertToInline.KeyTip = "LI"
        Me.grpReport_PlH_convertToInline.Label = "Force inline"
        Me.grpReport_PlH_convertToInline.Name = "grpReport_PlH_convertToInline"
        Me.grpReport_PlH_convertToInline.OfficeImageId = "TextWrappingInLineWithText"
        Me.grpReport_PlH_convertToInline.ScreenTip = "Will force the selected table to 'inline'"
        Me.grpReport_PlH_convertToInline.ShowImage = True
        Me.grpReport_PlH_convertToInline.SuperTip = "'Accessible' documents require Tab;es/Placeholdersto be 'regular' and inline. Use" &
    " the 'Placeholder Map' tool to find and convert all floating Placeholders to inl" &
    "ine."
        '
        'grp_Plh_miscPlaceholders
        '
        Me.grp_Plh_miscPlaceholders.Items.Add(Me.grpPicts_PasteAsPic)
        Me.grp_Plh_miscPlaceholders.Items.Add(Me.grpEquations_Numbered)
        Me.grp_Plh_miscPlaceholders.Label = "Miscellaneous Functions"
        Me.grp_Plh_miscPlaceholders.Name = "grp_Plh_miscPlaceholders"
        '
        'grpPicts_PasteAsPic
        '
        Me.grpPicts_PasteAsPic.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grpPicts_PasteAsPic.KeyTip = "PP"
        Me.grpPicts_PasteAsPic.Label = "Paste as pic"
        Me.grpPicts_PasteAsPic.Name = "grpPicts_PasteAsPic"
        Me.grpPicts_PasteAsPic.OfficeImageId = "ChartAreaChart"
        Me.grpPicts_PasteAsPic.ScreenTip = "Paste Clipboard contents as an image"
        Me.grpPicts_PasteAsPic.ShowImage = True
        Me.grpPicts_PasteAsPic.SuperTip = "Use this option to paste items such as an Excel chart into a placeholder as an im" &
    "age."
        '
        'grpEquations_Numbered
        '
        Me.grpEquations_Numbered.KeyTip = "EN"
        Me.grpEquations_Numbered.Label = "Numbered Equation"
        Me.grpEquations_Numbered.Name = "grpEquations_Numbered"
        Me.grpEquations_Numbered.OfficeImageId = "EquationOptions"
        Me.grpEquations_Numbered.ScreenTip = "Numbered Equation"
        Me.grpEquations_Numbered.ShowImage = True
        Me.grpEquations_Numbered.SuperTip = resources.GetString("grpEquations_Numbered.SuperTip")
        '
        'tab_aa_PagesAndSections
        '
        Me.tab_aa_PagesAndSections.Groups.Add(Me.grpRpt_CoversAndTOC)
        Me.tab_aa_PagesAndSections.Groups.Add(Me.grpRpt_ImagePanels)
        Me.tab_aa_PagesAndSections.Groups.Add(Me.grpRpt_Report)
        Me.tab_aa_PagesAndSections.Groups.Add(Me.grpRpt_Appendix)
        Me.tab_aa_PagesAndSections.Groups.Add(Me.grpRpt_sectOptions)
        Me.tab_aa_PagesAndSections.Groups.Add(Me.grpRpt_CoveringLetter)
        Me.tab_aa_PagesAndSections.Groups.Add(Me.grp_WhatsNew)
        Me.tab_aa_PagesAndSections.Groups.Add(Me.grp_Fixes)
        Me.tab_aa_PagesAndSections.KeyTip = "JP"
        Me.tab_aa_PagesAndSections.Label = "Pages and Sections"
        Me.tab_aa_PagesAndSections.Name = "tab_aa_PagesAndSections"
        Me.tab_aa_PagesAndSections.Position = Me.Factory.RibbonPosition.BeforeOfficeId("TabInsert")
        '
        'grpRpt_CoversAndTOC
        '
        Me.grpRpt_CoversAndTOC.Items.Add(Me.gal_CoverPages)
        Me.grpRpt_CoversAndTOC.Items.Add(Me.grpCntsPages)
        Me.grpRpt_CoversAndTOC.Items.Add(Me.grpCoversToc_mnu2)
        Me.grpRpt_CoversAndTOC.Label = "Covers and TOC"
        Me.grpRpt_CoversAndTOC.Name = "grpRpt_CoversAndTOC"
        '
        'gal_CoverPages
        '
        Me.gal_CoverPages.Buttons.Add(Me.gal_CoverPages_btn_deleteCoverPage)
        Me.gal_CoverPages.ColumnCount = 2
        RibbonDropDownItemImpl1.Image = Global.AA_GeneralReport_Addin.My.Resources.Resources.CP_TG_filledPattern_small
        RibbonDropDownItemImpl1.Label = "cp_00"
        RibbonDropDownItemImpl1.ScreenTip = "Filled Patttern"
        RibbonDropDownItemImpl1.Tag = "0-cp_TG_filledPattern"
        RibbonDropDownItemImpl2.Image = Global.AA_GeneralReport_Addin.My.Resources.Resources.Cp_TG_emptyPattern_small
        RibbonDropDownItemImpl2.Label = "cp_01"
        RibbonDropDownItemImpl2.ScreenTip = "Empty Pattern"
        RibbonDropDownItemImpl2.Tag = "1-cp_TG_emptyPattern"
        RibbonDropDownItemImpl3.Image = Global.AA_GeneralReport_Addin.My.Resources.Resources.Cp_TG_picturePattern_small
        RibbonDropDownItemImpl3.Label = "cp_02"
        RibbonDropDownItemImpl3.ScreenTip = "Picture with Pattern"
        RibbonDropDownItemImpl3.Tag = "2-cp_TG_picturePattern"
        Me.gal_CoverPages.Items.Add(RibbonDropDownItemImpl1)
        Me.gal_CoverPages.Items.Add(RibbonDropDownItemImpl2)
        Me.gal_CoverPages.Items.Add(RibbonDropDownItemImpl3)
        Me.gal_CoverPages.KeyTip = "CV"
        Me.gal_CoverPages.Label = "Cover Pages"
        Me.gal_CoverPages.Name = "gal_CoverPages"
        Me.gal_CoverPages.OfficeImageId = "CustomCoverPageGallery"
        Me.gal_CoverPages.ShowImage = True
        Me.gal_CoverPages.ShowItemLabel = False
        Me.gal_CoverPages.SuperTip = """Select an Acil Allen Cover Page"""
        '
        'gal_CoverPages_btn_deleteCoverPage
        '
        Me.gal_CoverPages_btn_deleteCoverPage.Label = "Delete Cover Page"
        Me.gal_CoverPages_btn_deleteCoverPage.Name = "gal_CoverPages_btn_deleteCoverPage"
        Me.gal_CoverPages_btn_deleteCoverPage.OfficeImageId = "BevelShapeGallery"
        Me.gal_CoverPages_btn_deleteCoverPage.ScreenTip = "Delete Cover Page"
        Me.gal_CoverPages_btn_deleteCoverPage.ShowImage = True
        Me.gal_CoverPages_btn_deleteCoverPage.SuperTip = "If there is a Cover Page anywhere in the document this function will delete it. Y" &
    "ou don't have to place your cursor in the Cover Page."
        '
        'grpCntsPages
        '
        Me.grpCntsPages.Items.Add(Me.grpContactsPages_FrontPage_AckOfCountry)
        Me.grpCntsPages.Items.Add(Me.grpContactsPages_FrontPage)
        Me.grpCntsPages.Items.Add(Me.Separator22)
        Me.grpCntsPages.Items.Add(Me.grpContactsPages_BackPage)
        Me.grpCntsPages.Items.Add(Me.Separator23)
        Me.grpCntsPages.Items.Add(Me.grpContactsPages_mnu_2)
        Me.grpCntsPages.Items.Add(Me.Separator24)
        Me.grpCntsPages.Items.Add(Me.grpContactsPages_CopyrightStatement)
        Me.grpCntsPages.Items.Add(Me.grpContactsPages_Disclaimer)
        Me.grpCntsPages.KeyTip = "CC"
        Me.grpCntsPages.Label = "Contacts Pages"
        Me.grpCntsPages.Name = "grpCntsPages"
        Me.grpCntsPages.OfficeImageId = "DocumentPanelTemplate"
        Me.grpCntsPages.ScreenTip = "Contacts Pages"
        Me.grpCntsPages.ShowImage = True
        Me.grpCntsPages.SuperTip = """Inserts Acil Allen contact information as a new section, either at the current c" &
    "ursor position or at the back page."""
        '
        'grpContactsPages_FrontPage_AckOfCountry
        '
        Me.grpContactsPages_FrontPage_AckOfCountry.Label = "Insert Front Contacts Page (&Acknowledgement of Country)"
        Me.grpContactsPages_FrontPage_AckOfCountry.Name = "grpContactsPages_FrontPage_AckOfCountry"
        Me.grpContactsPages_FrontPage_AckOfCountry.OfficeImageId = "BevelShapeGallery"
        Me.grpContactsPages_FrontPage_AckOfCountry.ScreenTip = "Insert Front Contacts Page (Acknowledgement of Country)"
        Me.grpContactsPages_FrontPage_AckOfCountry.ShowImage = True
        Me.grpContactsPages_FrontPage_AckOfCountry.SuperTip = """Inserts a new section at the current cursor position which contains office addre" &
    "ss blocks, logo and 'Acknowledgement of Country'."""
        '
        'grpContactsPages_FrontPage
        '
        Me.grpContactsPages_FrontPage.Label = "Insert Front Contacts Page (&Standard Text)"
        Me.grpContactsPages_FrontPage.Name = "grpContactsPages_FrontPage"
        Me.grpContactsPages_FrontPage.OfficeImageId = "BevelShapeGallery"
        Me.grpContactsPages_FrontPage.ScreenTip = "Insert Front Contacts Page (Standard Text)"
        Me.grpContactsPages_FrontPage.ShowImage = True
        Me.grpContactsPages_FrontPage.SuperTip = """Inserts a new section at the current cursor position which contains office addre" &
    "ss blocks and logo."""
        '
        'Separator22
        '
        Me.Separator22.Name = "Separator22"
        '
        'grpContactsPages_BackPage
        '
        Me.grpContactsPages_BackPage.Label = "Insert &Back Contacts Page"
        Me.grpContactsPages_BackPage.Name = "grpContactsPages_BackPage"
        Me.grpContactsPages_BackPage.OfficeImageId = "BevelShapeGallery"
        Me.grpContactsPages_BackPage.ScreenTip = "Insert Back Contacts Page"
        Me.grpContactsPages_BackPage.ShowImage = True
        Me.grpContactsPages_BackPage.SuperTip = "Inserts a new section at the end of the document which contains office address bl" &
    "ocks and logo."
        '
        'Separator23
        '
        Me.Separator23.Name = "Separator23"
        '
        'grpContactsPages_mnu_2
        '
        Me.grpContactsPages_mnu_2.Items.Add(Me.grpContactsPages_ReportTo)
        Me.grpContactsPages_mnu_2.Items.Add(Me.grpContactsPages_ProposalTo)
        Me.grpContactsPages_mnu_2.Label = "Report &type options (Front Contacts Page)"
        Me.grpContactsPages_mnu_2.Name = "grpContactsPages_mnu_2"
        Me.grpContactsPages_mnu_2.OfficeImageId = "BevelShapeGallery"
        Me.grpContactsPages_mnu_2.ScreenTip = "Report type options (Front Contacts Page)"
        Me.grpContactsPages_mnu_2.ShowImage = True
        Me.grpContactsPages_mnu_2.SuperTip = """Allows the user to select document type options such as 'Report to', 'Proposal t" &
    "o' etc"""
        '
        'grpContactsPages_ReportTo
        '
        Me.grpContactsPages_ReportTo.Label = "Insert '&Report to' option"
        Me.grpContactsPages_ReportTo.Name = "grpContactsPages_ReportTo"
        Me.grpContactsPages_ReportTo.OfficeImageId = "BevelShapeGallery"
        Me.grpContactsPages_ReportTo.ScreenTip = "Insert 'Report to' option"
        Me.grpContactsPages_ReportTo.ShowImage = True
        Me.grpContactsPages_ReportTo.SuperTip = """Inserts the 'Report to' option into the Front Contacts Page."""
        '
        'grpContactsPages_ProposalTo
        '
        Me.grpContactsPages_ProposalTo.Label = "Insert '&Proposal to' option"
        Me.grpContactsPages_ProposalTo.Name = "grpContactsPages_ProposalTo"
        Me.grpContactsPages_ProposalTo.OfficeImageId = "BevelShapeGallery"
        Me.grpContactsPages_ProposalTo.ScreenTip = "Insert 'Proposal to' option"
        Me.grpContactsPages_ProposalTo.ShowImage = True
        Me.grpContactsPages_ProposalTo.SuperTip = """Inserts the 'Proposal to' option into the Front Contacts Page."""
        '
        'Separator24
        '
        Me.Separator24.Name = "Separator24"
        '
        'grpContactsPages_CopyrightStatement
        '
        Me.grpContactsPages_CopyrightStatement.Label = "Insert Acil Allen &Copyright Statement (Front Contacts Page)"
        Me.grpContactsPages_CopyrightStatement.Name = "grpContactsPages_CopyrightStatement"
        Me.grpContactsPages_CopyrightStatement.OfficeImageId = "BevelShapeGallery"
        Me.grpContactsPages_CopyrightStatement.ScreenTip = "Insert Acil Allen Copyright Statement (Front Contacts Page)"
        Me.grpContactsPages_CopyrightStatement.ShowImage = True
        Me.grpContactsPages_CopyrightStatement.SuperTip = """Inserts the Acil Allen copyright statement into the Front Contacts Page."""
        '
        'grpContactsPages_Disclaimer
        '
        Me.grpContactsPages_Disclaimer.Label = "Insert Acil Allen &Disclaimer Statement (Front Contacts Page)"
        Me.grpContactsPages_Disclaimer.Name = "grpContactsPages_Disclaimer"
        Me.grpContactsPages_Disclaimer.OfficeImageId = "BevelShapeGallery"
        Me.grpContactsPages_Disclaimer.ScreenTip = "Insert Acil Allen Disclaimer Statement (Front Contacts Page)"
        Me.grpContactsPages_Disclaimer.ShowImage = True
        Me.grpContactsPages_Disclaimer.SuperTip = """Inserts the Acil Allen disclaimer statement into the Front Contacts Page."""
        '
        'grpCoversToc_mnu2
        '
        Me.grpCoversToc_mnu2.Items.Add(Me.grpToc_TOC_insertSection)
        Me.grpCoversToc_mnu2.Items.Add(Me.Separator20)
        Me.grpCoversToc_mnu2.Items.Add(Me.grpToc_TOC_insertLevels_1_to_1)
        Me.grpCoversToc_mnu2.Items.Add(Me.grpToc_TOC_insertLevels_1_to_2)
        Me.grpCoversToc_mnu2.Items.Add(Me.grpToc_TOC_insertLevels_1_to_3)
        Me.grpCoversToc_mnu2.Items.Add(Me.Separator21)
        Me.grpCoversToc_mnu2.Items.Add(Me.grpToc_TOC_update)
        Me.grpCoversToc_mnu2.KeyTip = "CT"
        Me.grpCoversToc_mnu2.Label = "TOC Functions"
        Me.grpCoversToc_mnu2.Name = "grpCoversToc_mnu2"
        Me.grpCoversToc_mnu2.OfficeImageId = "TableOfContentsGallery"
        Me.grpCoversToc_mnu2.ScreenTip = "TOC Functions"
        Me.grpCoversToc_mnu2.ShowImage = True
        Me.grpCoversToc_mnu2.SuperTip = resources.GetString("grpCoversToc_mnu2.SuperTip")
        '
        'grpToc_TOC_insertSection
        '
        Me.grpToc_TOC_insertSection.Label = "&Insert a TOC section"
        Me.grpToc_TOC_insertSection.Name = "grpToc_TOC_insertSection"
        Me.grpToc_TOC_insertSection.OfficeImageId = "BevelShapeGallery"
        Me.grpToc_TOC_insertSection.ScreenTip = "Insert a TOC section"
        Me.grpToc_TOC_insertSection.ShowImage = True
        Me.grpToc_TOC_insertSection.SuperTip = """Inserts a Table of Contents section with a TOC."""
        '
        'Separator20
        '
        Me.Separator20.Name = "Separator20"
        '
        'grpToc_TOC_insertLevels_1_to_1
        '
        Me.grpToc_TOC_insertLevels_1_to_1.Label = "TOC to show &1 level"
        Me.grpToc_TOC_insertLevels_1_to_1.Name = "grpToc_TOC_insertLevels_1_to_1"
        Me.grpToc_TOC_insertLevels_1_to_1.OfficeImageId = "BevelShapeGallery"
        Me.grpToc_TOC_insertLevels_1_to_1.ScreenTip = "TOC to show 1 level"
        Me.grpToc_TOC_insertLevels_1_to_1.ShowImage = True
        Me.grpToc_TOC_insertLevels_1_to_1.SuperTip = """Replaces the contents of an existing TOC section with a TOC field formatted to s" &
    "how a 1 level TOC.... Make certain your cursor is in the existing TOC."""
        '
        'grpToc_TOC_insertLevels_1_to_2
        '
        Me.grpToc_TOC_insertLevels_1_to_2.Label = "TOC to show &2 levels"
        Me.grpToc_TOC_insertLevels_1_to_2.Name = "grpToc_TOC_insertLevels_1_to_2"
        Me.grpToc_TOC_insertLevels_1_to_2.OfficeImageId = "BevelShapeGallery"
        Me.grpToc_TOC_insertLevels_1_to_2.ScreenTip = "TOC to show 2 levels"
        Me.grpToc_TOC_insertLevels_1_to_2.ShowImage = True
        Me.grpToc_TOC_insertLevels_1_to_2.SuperTip = """Replaces the contents of an existing TOC section with a TOC field formatted to s" &
    "how a 2 level TOC.... Make certain your cursor is in the existing TOC."""
        '
        'grpToc_TOC_insertLevels_1_to_3
        '
        Me.grpToc_TOC_insertLevels_1_to_3.Label = "TOC to show &3 levels"
        Me.grpToc_TOC_insertLevels_1_to_3.Name = "grpToc_TOC_insertLevels_1_to_3"
        Me.grpToc_TOC_insertLevels_1_to_3.OfficeImageId = "BevelShapeGallery"
        Me.grpToc_TOC_insertLevels_1_to_3.ScreenTip = "TOC to show 3 levels"
        Me.grpToc_TOC_insertLevels_1_to_3.ShowImage = True
        Me.grpToc_TOC_insertLevels_1_to_3.SuperTip = """Replaces the contents of an existing TOC section with a TOC field formatted to s" &
    "how a 3 level TOC.... Make certain your cursor is in the existing TOC."""
        '
        'Separator21
        '
        Me.Separator21.Name = "Separator21"
        '
        'grpToc_TOC_update
        '
        Me.grpToc_TOC_update.Label = "&Update TOC"
        Me.grpToc_TOC_update.Name = "grpToc_TOC_update"
        Me.grpToc_TOC_update.OfficeImageId = "BevelShapeGallery"
        Me.grpToc_TOC_update.ScreenTip = "Update TOC"
        Me.grpToc_TOC_update.ShowImage = True
        '
        'grpRpt_ImagePanels
        '
        Me.grpRpt_ImagePanels.Items.Add(Me.grpCoversToc_mnu_Images)
        Me.grpRpt_ImagePanels.Items.Add(Me.grpImageHandling_mnu_ImgSection)
        Me.grpRpt_ImagePanels.Items.Add(Me.grpImageHandling_mnu_FillBackPanel)
        Me.grpRpt_ImagePanels.Label = "Image Panels"
        Me.grpRpt_ImagePanels.Name = "grpRpt_ImagePanels"
        '
        'grpCoversToc_mnu_Images
        '
        Me.grpCoversToc_mnu_Images.Items.Add(Me.grpCpImages_ImageFromFile)
        Me.grpCoversToc_mnu_Images.Items.Add(Me.grpCpImages_ImageFromClip)
        Me.grpCoversToc_mnu_Images.Items.Add(Me.Separator26)
        Me.grpCoversToc_mnu_Images.Items.Add(Me.grpCpImages_BackPanelFill_RawImageFromFile)
        Me.grpCoversToc_mnu_Images.Items.Add(Me.Separator25)
        Me.grpCoversToc_mnu_Images.Items.Add(Me.grpCpImages_Delete_SmallPict)
        Me.grpCoversToc_mnu_Images.KeyTip = "CP"
        Me.grpCoversToc_mnu_Images.Label = "Small Cover Page Picture"
        Me.grpCoversToc_mnu_Images.Name = "grpCoversToc_mnu_Images"
        Me.grpCoversToc_mnu_Images.OfficeImageId = "PictureInsertFromFilePowerPoint"
        Me.grpCoversToc_mnu_Images.ScreenTip = "Small Cover Page Picture"
        Me.grpCoversToc_mnu_Images.ShowImage = True
        Me.grpCoversToc_mnu_Images.SuperTip = resources.GetString("grpCoversToc_mnu_Images.SuperTip")
        '
        'grpCpImages_ImageFromFile
        '
        Me.grpCpImages_ImageFromFile.Label = "Replace with cropped Image from &File..."
        Me.grpCpImages_ImageFromFile.Name = "grpCpImages_ImageFromFile"
        Me.grpCpImages_ImageFromFile.OfficeImageId = "PictureCrop"
        Me.grpCpImages_ImageFromFile.ScreenTip = "From File"
        Me.grpCpImages_ImageFromFile.ShowImage = True
        Me.grpCpImages_ImageFromFile.SuperTip = """Insert a cropped image from a file."""
        '
        'grpCpImages_ImageFromClip
        '
        Me.grpCpImages_ImageFromClip.Label = "Replace with cropped Image from the &Clipboard..."
        Me.grpCpImages_ImageFromClip.Name = "grpCpImages_ImageFromClip"
        Me.grpCpImages_ImageFromClip.OfficeImageId = "PictureCrop"
        Me.grpCpImages_ImageFromClip.ScreenTip = "From Clipboard"
        Me.grpCpImages_ImageFromClip.ShowImage = True
        Me.grpCpImages_ImageFromClip.SuperTip = "Insert cropped custom image from the clipboard."
        '
        'Separator26
        '
        Me.Separator26.Name = "Separator26"
        '
        'grpCpImages_BackPanelFill_RawImageFromFile
        '
        Me.grpCpImages_BackPanelFill_RawImageFromFile.Label = "&Raw Image from File..."
        Me.grpCpImages_BackPanelFill_RawImageFromFile.Name = "grpCpImages_BackPanelFill_RawImageFromFile"
        Me.grpCpImages_BackPanelFill_RawImageFromFile.OfficeImageId = "PictureCrop"
        Me.grpCpImages_BackPanelFill_RawImageFromFile.ScreenTip = "Raw Image from File"
        Me.grpCpImages_BackPanelFill_RawImageFromFile.ShowImage = True
        Me.grpCpImages_BackPanelFill_RawImageFromFile.SuperTip = """Insert a raw image from a file. Note that the image must have the same aspect ra" &
    "tio (height/width) as the page you are working on. Otherwise the result will be " &
    "distorted"""
        '
        'Separator25
        '
        Me.Separator25.Name = "Separator25"
        '
        'grpCpImages_Delete_SmallPict
        '
        Me.grpCpImages_Delete_SmallPict.Label = "&Delete"
        Me.grpCpImages_Delete_SmallPict.Name = "grpCpImages_Delete_SmallPict"
        Me.grpCpImages_Delete_SmallPict.OfficeImageId = "PictureCrop"
        Me.grpCpImages_Delete_SmallPict.ScreenTip = "Delete"
        Me.grpCpImages_Delete_SmallPict.ShowImage = True
        Me.grpCpImages_Delete_SmallPict.SuperTip = """Will delete the small picture"""
        '
        'grpImageHandling_mnu_ImgSection
        '
        Me.grpImageHandling_mnu_ImgSection.Items.Add(Me.grpImageHandling_insert_BackPanel)
        Me.grpImageHandling_mnu_ImgSection.Items.Add(Me.Separator27)
        Me.grpImageHandling_mnu_ImgSection.Items.Add(Me.grpImageHandling_delete_BackPanel)
        Me.grpImageHandling_mnu_ImgSection.Label = "Insert image back panel"
        Me.grpImageHandling_mnu_ImgSection.Name = "grpImageHandling_mnu_ImgSection"
        Me.grpImageHandling_mnu_ImgSection.OfficeImageId = "PictureInsertFromFilePowerPoint"
        Me.grpImageHandling_mnu_ImgSection.ScreenTip = "Insert image back panel"
        Me.grpImageHandling_mnu_ImgSection.ShowImage = True
        Me.grpImageHandling_mnu_ImgSection.SuperTip = """Use a menu item to insert an image back panel (in the current section) or a sect" &
    "ion containing an image back panel at the current cursor position"""
        '
        'grpImageHandling_insert_BackPanel
        '
        Me.grpImageHandling_insert_BackPanel.Label = "&Insert image back panel"
        Me.grpImageHandling_insert_BackPanel.Name = "grpImageHandling_insert_BackPanel"
        Me.grpImageHandling_insert_BackPanel.OfficeImageId = "FileNew"
        Me.grpImageHandling_insert_BackPanel.ScreenTip = "Insert image back panel"
        Me.grpImageHandling_insert_BackPanel.ShowImage = True
        Me.grpImageHandling_insert_BackPanel.SuperTip = """This menu item inserts an image back panel in the current section."""
        '
        'Separator27
        '
        Me.Separator27.Name = "Separator27"
        '
        'grpImageHandling_delete_BackPanel
        '
        Me.grpImageHandling_delete_BackPanel.Label = "&Delete image back panel"
        Me.grpImageHandling_delete_BackPanel.Name = "grpImageHandling_delete_BackPanel"
        Me.grpImageHandling_delete_BackPanel.OfficeImageId = "FileNew"
        Me.grpImageHandling_delete_BackPanel.ScreenTip = "Delete image back panel"
        Me.grpImageHandling_delete_BackPanel.ShowImage = True
        Me.grpImageHandling_delete_BackPanel.SuperTip = """This menu item will delete an existing back panel in the current section."""
        '
        'grpImageHandling_mnu_FillBackPanel
        '
        Me.grpImageHandling_mnu_FillBackPanel.Items.Add(Me.grpImageHandling_BackPanelFill_FromFile)
        Me.grpImageHandling_mnu_FillBackPanel.Items.Add(Me.grpImageHandling_BackPanelFill_FromClipBoard)
        Me.grpImageHandling_mnu_FillBackPanel.Items.Add(Me.Separator28)
        Me.grpImageHandling_mnu_FillBackPanel.Items.Add(Me.grpImageHandling_BackPanelFill_RawImageFromFile)
        Me.grpImageHandling_mnu_FillBackPanel.Items.Add(Me.Separator29)
        Me.grpImageHandling_mnu_FillBackPanel.Items.Add(Me.grpImageHandling_Reset_backcolour)
        Me.grpImageHandling_mnu_FillBackPanel.Items.Add(Me.grpImageHandling_Reset_backcolour_to_CaseStudyGrey)
        Me.grpImageHandling_mnu_FillBackPanel.Items.Add(Me.grpImageHandling_Custom_backcolour)
        Me.grpImageHandling_mnu_FillBackPanel.Items.Add(Me.Separator30)
        Me.grpImageHandling_mnu_FillBackPanel.Items.Add(Me.grpImageHandling_submnu_FillBackPanel_SetTransparency)
        Me.grpImageHandling_mnu_FillBackPanel.Items.Add(Me.mnu_SetBackPanel_to_BannerHeight)
        Me.grpImageHandling_mnu_FillBackPanel.KeyTip = "BC"
        Me.grpImageHandling_mnu_FillBackPanel.Label = "Set Image back panel to..."
        Me.grpImageHandling_mnu_FillBackPanel.Name = "grpImageHandling_mnu_FillBackPanel"
        Me.grpImageHandling_mnu_FillBackPanel.OfficeImageId = "PictureInsertFromFilePowerPoint"
        Me.grpImageHandling_mnu_FillBackPanel.ScreenTip = "Set Image back panel to..."
        Me.grpImageHandling_mnu_FillBackPanel.ShowImage = True
        Me.grpImageHandling_mnu_FillBackPanel.SuperTip = resources.GetString("grpImageHandling_mnu_FillBackPanel.SuperTip")
        '
        'grpImageHandling_BackPanelFill_FromFile
        '
        Me.grpImageHandling_BackPanelFill_FromFile.Label = "Cropped Image from &File..."
        Me.grpImageHandling_BackPanelFill_FromFile.Name = "grpImageHandling_BackPanelFill_FromFile"
        Me.grpImageHandling_BackPanelFill_FromFile.OfficeImageId = "PictureCrop"
        Me.grpImageHandling_BackPanelFill_FromFile.ScreenTip = "Cropped Image from File"
        Me.grpImageHandling_BackPanelFill_FromFile.ShowImage = True
        Me.grpImageHandling_BackPanelFill_FromFile.SuperTip = """Insert a cropped image from a file."""
        '
        'grpImageHandling_BackPanelFill_FromClipBoard
        '
        Me.grpImageHandling_BackPanelFill_FromClipBoard.Label = "Cropped Image from the &Clipboard..."
        Me.grpImageHandling_BackPanelFill_FromClipBoard.Name = "grpImageHandling_BackPanelFill_FromClipBoard"
        Me.grpImageHandling_BackPanelFill_FromClipBoard.OfficeImageId = "PictureCrop"
        Me.grpImageHandling_BackPanelFill_FromClipBoard.ScreenTip = "Cropped Image from the Clipboard"
        Me.grpImageHandling_BackPanelFill_FromClipBoard.ShowImage = True
        Me.grpImageHandling_BackPanelFill_FromClipBoard.SuperTip = """Insert cropped custom image from the clipboard."""
        '
        'Separator28
        '
        Me.Separator28.Name = "Separator28"
        '
        'grpImageHandling_BackPanelFill_RawImageFromFile
        '
        Me.grpImageHandling_BackPanelFill_RawImageFromFile.Label = "Raw Image from File..."
        Me.grpImageHandling_BackPanelFill_RawImageFromFile.Name = "grpImageHandling_BackPanelFill_RawImageFromFile"
        Me.grpImageHandling_BackPanelFill_RawImageFromFile.OfficeImageId = "PictureCrop"
        Me.grpImageHandling_BackPanelFill_RawImageFromFile.ScreenTip = "Raw Image from File."
        Me.grpImageHandling_BackPanelFill_RawImageFromFile.ShowImage = True
        Me.grpImageHandling_BackPanelFill_RawImageFromFile.SuperTip = """Insert a raw image from a file. Note that the image must have the same aspect ra" &
    "tio (height/width) as the page you are working on. Otherwise the result will be " &
    "distorted"""
        '
        'Separator29
        '
        Me.Separator29.Name = "Separator29"
        '
        'grpImageHandling_Reset_backcolour
        '
        Me.grpImageHandling_Reset_backcolour.Label = "Standard purple colour"
        Me.grpImageHandling_Reset_backcolour.Name = "grpImageHandling_Reset_backcolour"
        Me.grpImageHandling_Reset_backcolour.OfficeImageId = "PictureCrop"
        Me.grpImageHandling_Reset_backcolour.ScreenTip = "Standard purple colour"
        Me.grpImageHandling_Reset_backcolour.ShowImage = True
        Me.grpImageHandling_Reset_backcolour.SuperTip = """Will fill the back panel with the default AA purple"""
        '
        'grpImageHandling_Reset_backcolour_to_CaseStudyGrey
        '
        Me.grpImageHandling_Reset_backcolour_to_CaseStudyGrey.Label = "Case Study grey"
        Me.grpImageHandling_Reset_backcolour_to_CaseStudyGrey.Name = "grpImageHandling_Reset_backcolour_to_CaseStudyGrey"
        Me.grpImageHandling_Reset_backcolour_to_CaseStudyGrey.OfficeImageId = "PictureCrop"
        Me.grpImageHandling_Reset_backcolour_to_CaseStudyGrey.ScreenTip = "Case Study grey"
        Me.grpImageHandling_Reset_backcolour_to_CaseStudyGrey.ShowImage = True
        Me.grpImageHandling_Reset_backcolour_to_CaseStudyGrey.SuperTip = """Will fill the back panel with the standard grey used for case studies"""
        '
        'grpImageHandling_Custom_backcolour
        '
        Me.grpImageHandling_Custom_backcolour.Label = "Custom colour"
        Me.grpImageHandling_Custom_backcolour.Name = "grpImageHandling_Custom_backcolour"
        Me.grpImageHandling_Custom_backcolour.OfficeImageId = "ViewBackToColorView"
        Me.grpImageHandling_Custom_backcolour.ScreenTip = "Custom colour"
        Me.grpImageHandling_Custom_backcolour.ShowImage = True
        Me.grpImageHandling_Custom_backcolour.SuperTip = """Will fill the back panel with a custom colour selected from a custom colour dial" &
    "og"""
        '
        'Separator30
        '
        Me.Separator30.Name = "Separator30"
        '
        'grpImageHandling_submnu_FillBackPanel_SetTransparency
        '
        Me.grpImageHandling_submnu_FillBackPanel_SetTransparency.Items.Add(Me.submnu_SetTransparency_to_0)
        Me.grpImageHandling_submnu_FillBackPanel_SetTransparency.Items.Add(Me.submnu_SetTransparency_to_25)
        Me.grpImageHandling_submnu_FillBackPanel_SetTransparency.Items.Add(Me.submnu_SetTransparency_to_50)
        Me.grpImageHandling_submnu_FillBackPanel_SetTransparency.Items.Add(Me.submnu_SetTransparency_to_75)
        Me.grpImageHandling_submnu_FillBackPanel_SetTransparency.Items.Add(Me.submnu_SetTransparency_to_100)
        Me.grpImageHandling_submnu_FillBackPanel_SetTransparency.Label = "Set &Transparency..."
        Me.grpImageHandling_submnu_FillBackPanel_SetTransparency.Name = "grpImageHandling_submnu_FillBackPanel_SetTransparency"
        Me.grpImageHandling_submnu_FillBackPanel_SetTransparency.OfficeImageId = "PictureInsertFromFilePowerPoint"
        Me.grpImageHandling_submnu_FillBackPanel_SetTransparency.ScreenTip = "Transparency"
        Me.grpImageHandling_submnu_FillBackPanel_SetTransparency.ShowImage = True
        Me.grpImageHandling_submnu_FillBackPanel_SetTransparency.SuperTip = resources.GetString("grpImageHandling_submnu_FillBackPanel_SetTransparency.SuperTip")
        '
        'submnu_SetTransparency_to_0
        '
        Me.submnu_SetTransparency_to_0.Label = "&0% transparent (fully opaque)"
        Me.submnu_SetTransparency_to_0.Name = "submnu_SetTransparency_to_0"
        Me.submnu_SetTransparency_to_0.OfficeImageId = "BevelShapeGallery"
        Me.submnu_SetTransparency_to_0.ScreenTip = "Fully opaque"
        Me.submnu_SetTransparency_to_0.ShowImage = True
        Me.submnu_SetTransparency_to_0.SuperTip = """Will set the image back panel to 0% transparent"""
        '
        'submnu_SetTransparency_to_25
        '
        Me.submnu_SetTransparency_to_25.Label = "&25% transparent"
        Me.submnu_SetTransparency_to_25.Name = "submnu_SetTransparency_to_25"
        Me.submnu_SetTransparency_to_25.OfficeImageId = "BevelShapeGallery"
        Me.submnu_SetTransparency_to_25.ScreenTip = "25%"
        Me.submnu_SetTransparency_to_25.ShowImage = True
        Me.submnu_SetTransparency_to_25.SuperTip = """Will set the image back panel to 25% transparent"""
        '
        'submnu_SetTransparency_to_50
        '
        Me.submnu_SetTransparency_to_50.Label = "&50% transparent"
        Me.submnu_SetTransparency_to_50.Name = "submnu_SetTransparency_to_50"
        Me.submnu_SetTransparency_to_50.OfficeImageId = "BevelShapeGallery"
        Me.submnu_SetTransparency_to_50.ScreenTip = "50%"
        Me.submnu_SetTransparency_to_50.ShowImage = True
        Me.submnu_SetTransparency_to_50.SuperTip = """Will set the image back panel to 50% transparent"""
        '
        'submnu_SetTransparency_to_75
        '
        Me.submnu_SetTransparency_to_75.Label = "&75% transparent"
        Me.submnu_SetTransparency_to_75.Name = "submnu_SetTransparency_to_75"
        Me.submnu_SetTransparency_to_75.OfficeImageId = "BevelShapeGallery"
        Me.submnu_SetTransparency_to_75.ScreenTip = "75%"
        Me.submnu_SetTransparency_to_75.ShowImage = True
        Me.submnu_SetTransparency_to_75.SuperTip = """Will set the image back panel to 75% transparent"""
        '
        'submnu_SetTransparency_to_100
        '
        Me.submnu_SetTransparency_to_100.Label = "&100% transparent"
        Me.submnu_SetTransparency_to_100.Name = "submnu_SetTransparency_to_100"
        Me.submnu_SetTransparency_to_100.OfficeImageId = "BevelShapeGallery"
        Me.submnu_SetTransparency_to_100.ScreenTip = "Fully Transparent"
        Me.submnu_SetTransparency_to_100.ShowImage = True
        Me.submnu_SetTransparency_to_100.SuperTip = """Will set the image back panel to 100% transparent"""
        '
        'mnu_SetBackPanel_to_BannerHeight
        '
        Me.mnu_SetBackPanel_to_BannerHeight.Label = "Set Back Panel height to 'Brief' banner height"
        Me.mnu_SetBackPanel_to_BannerHeight.Name = "mnu_SetBackPanel_to_BannerHeight"
        Me.mnu_SetBackPanel_to_BannerHeight.OfficeImageId = "PictureCrop"
        Me.mnu_SetBackPanel_to_BannerHeight.ScreenTip = "Brief banner to page height"
        Me.mnu_SetBackPanel_to_BannerHeight.ShowImage = True
        Me.mnu_SetBackPanel_to_BannerHeight.SuperTip = """Will set the image back panel height to the banner height used in the 'AA Brief'" &
    "."""
        '
        'grpRpt_Report
        '
        Me.grpRpt_Report.Items.Add(Me.grpRpt_btn_GlossaryAndAbbreviations_bblk)
        Me.grpRpt_Report.Items.Add(Me.grpReport_btn_newDivider_Chpt_bblk)
        Me.grpRpt_Report.Items.Add(Me.grpRpt_mnu_CreateExecSummary)
        Me.grpRpt_Report.Items.Add(Me.Separator2)
        Me.grpRpt_Report.Items.Add(Me.grpRpt_mnu_CreateRpt)
        Me.grpRpt_Report.Items.Add(Me.grpRpt_mnu_NewChapter)
        Me.grpRpt_Report.Items.Add(Me.grpRpt_mnu_Bibliography)
        Me.grpRpt_Report.Items.Add(Me.grpReport_btn_ToggleView)
        Me.grpRpt_Report.Items.Add(Me.grpRpt_mnu_RefreshDocument)
        Me.grpRpt_Report.Items.Add(Me.grpRpt_mnu_ApplyColour)
        Me.grpRpt_Report.Label = "Report"
        Me.grpRpt_Report.Name = "grpRpt_Report"
        '
        'grpRpt_btn_GlossaryAndAbbreviations_bblk
        '
        Me.grpRpt_btn_GlossaryAndAbbreviations_bblk.KeyTip = "RG"
        Me.grpRpt_btn_GlossaryAndAbbreviations_bblk.Label = "Glossary/Abbrev"
        Me.grpRpt_btn_GlossaryAndAbbreviations_bblk.Name = "grpRpt_btn_GlossaryAndAbbreviations_bblk"
        Me.grpRpt_btn_GlossaryAndAbbreviations_bblk.OfficeImageId = "PivotClearCustomOrdering"
        Me.grpRpt_btn_GlossaryAndAbbreviations_bblk.ShowImage = True
        '
        'grpReport_btn_newDivider_Chpt_bblk
        '
        Me.grpReport_btn_newDivider_Chpt_bblk.Image = Global.AA_GeneralReport_Addin.My.Resources.Resources.NewPart_TG
        Me.grpReport_btn_newDivider_Chpt_bblk.Label = "Part Divider"
        Me.grpReport_btn_newDivider_Chpt_bblk.Name = "grpReport_btn_newDivider_Chpt_bblk"
        Me.grpReport_btn_newDivider_Chpt_bblk.ShowImage = True
        '
        'grpRpt_mnu_CreateExecSummary
        '
        Me.grpRpt_mnu_CreateExecSummary.Items.Add(Me.grpExecSum_ExecSum_bblk)
        Me.grpRpt_mnu_CreateExecSummary.Items.Add(Me.grpExecSum_ExecSum_Grey_bblk)
        Me.grpRpt_mnu_CreateExecSummary.KeyTip = "RE"
        Me.grpRpt_mnu_CreateExecSummary.Label = "Exec Summary"
        Me.grpRpt_mnu_CreateExecSummary.Name = "grpRpt_mnu_CreateExecSummary"
        Me.grpRpt_mnu_CreateExecSummary.OfficeImageId = "SummarizeSlide"
        Me.grpRpt_mnu_CreateExecSummary.ShowImage = True
        Me.grpRpt_mnu_CreateExecSummary.SuperTip = """Inserts a new Excutive Summary"""
        '
        'grpExecSum_ExecSum_bblk
        '
        Me.grpExecSum_ExecSum_bblk.Label = "&Exec Summary"
        Me.grpExecSum_ExecSum_bblk.Name = "grpExecSum_ExecSum_bblk"
        Me.grpExecSum_ExecSum_bblk.OfficeImageId = "SummarizeSlide"
        Me.grpExecSum_ExecSum_bblk.ShowImage = True
        Me.grpExecSum_ExecSum_bblk.SuperTip = """Inserts a formatted Executive Summary section container."""
        '
        'grpExecSum_ExecSum_Grey_bblk
        '
        Me.grpExecSum_ExecSum_Grey_bblk.Label = "Exec Summary (&Grey)"
        Me.grpExecSum_ExecSum_Grey_bblk.Name = "grpExecSum_ExecSum_Grey_bblk"
        Me.grpExecSum_ExecSum_Grey_bblk.OfficeImageId = "SummarizeSlide"
        Me.grpExecSum_ExecSum_Grey_bblk.ShowImage = True
        Me.grpExecSum_ExecSum_Grey_bblk.SuperTip = """Inserts a formatted Executive Summary section with an image back panel set to rg" &
    "b(200, 200, 200)."""
        '
        'Separator2
        '
        Me.Separator2.Name = "Separator2"
        '
        'grpRpt_mnu_CreateRpt
        '
        Me.grpRpt_mnu_CreateRpt.Items.Add(Me.grpReport_btn_buildPortraitReport)
        Me.grpRpt_mnu_CreateRpt.Items.Add(Me.Separator3)
        Me.grpRpt_mnu_CreateRpt.Items.Add(Me.grpReprt_btn_buildLandscapeReport)
        Me.grpRpt_mnu_CreateRpt.Items.Add(Me.Separator4)
        Me.grpRpt_mnu_CreateRpt.Items.Add(Me.grpReport_btn_buildAABrief)
        Me.grpRpt_mnu_CreateRpt.KeyTip = "RN"
        Me.grpRpt_mnu_CreateRpt.Label = "Create Report or Brief"
        Me.grpRpt_mnu_CreateRpt.Name = "grpRpt_mnu_CreateRpt"
        Me.grpRpt_mnu_CreateRpt.OfficeImageId = "AnimationTriggerAddMenu"
        Me.grpRpt_mnu_CreateRpt.ShowImage = True
        Me.grpRpt_mnu_CreateRpt.Visible = False
        '
        'grpReport_btn_buildPortraitReport
        '
        Me.grpReport_btn_buildPortraitReport.Label = "Create a new &Portrait Report"
        Me.grpReport_btn_buildPortraitReport.Name = "grpReport_btn_buildPortraitReport"
        Me.grpReport_btn_buildPortraitReport.OfficeImageId = "SizeToGridAccess"
        Me.grpReport_btn_buildPortraitReport.ShowImage = True
        Me.grpReport_btn_buildPortraitReport.SuperTip = "Will create a new standard ACIL Allen portrait report skeleton"
        '
        'Separator3
        '
        Me.Separator3.Name = "Separator3"
        '
        'grpReprt_btn_buildLandscapeReport
        '
        Me.grpReprt_btn_buildLandscapeReport.Label = "Create a new &Landscape Report"
        Me.grpReprt_btn_buildLandscapeReport.Name = "grpReprt_btn_buildLandscapeReport"
        Me.grpReprt_btn_buildLandscapeReport.OfficeImageId = "SizeToGridAccess"
        Me.grpReprt_btn_buildLandscapeReport.ShowImage = True
        Me.grpReprt_btn_buildLandscapeReport.SuperTip = "Will create a new standard ACIL Allen landscape report skeleton"
        '
        'Separator4
        '
        Me.Separator4.Name = "Separator4"
        '
        'grpReport_btn_buildAABrief
        '
        Me.grpReport_btn_buildAABrief.Label = "Create a new ACIL Allen &Brief"
        Me.grpReport_btn_buildAABrief.Name = "grpReport_btn_buildAABrief"
        Me.grpReport_btn_buildAABrief.OfficeImageId = "SizeToGridAccess"
        Me.grpReport_btn_buildAABrief.ShowImage = True
        '
        'grpRpt_mnu_NewChapter
        '
        Me.grpRpt_mnu_NewChapter.Items.Add(Me.grpRpt_mnu_btn_NewChapter_inFront_bblk)
        Me.grpRpt_mnu_NewChapter.Items.Add(Me.grpRpt_mnu_btn_NewChapter_behind_bblk)
        Me.grpRpt_mnu_NewChapter.KeyTip = "RC"
        Me.grpRpt_mnu_NewChapter.Label = "New Chapter"
        Me.grpRpt_mnu_NewChapter.Name = "grpRpt_mnu_NewChapter"
        Me.grpRpt_mnu_NewChapter.OfficeImageId = "CompareAndCombine"
        Me.grpRpt_mnu_NewChapter.ScreenTip = "Inserts a new Chapter"
        Me.grpRpt_mnu_NewChapter.ShowImage = True
        '
        'grpRpt_mnu_btn_NewChapter_inFront_bblk
        '
        Me.grpRpt_mnu_btn_NewChapter_inFront_bblk.Label = "New Chapter (in &Front)"
        Me.grpRpt_mnu_btn_NewChapter_inFront_bblk.Name = "grpRpt_mnu_btn_NewChapter_inFront_bblk"
        Me.grpRpt_mnu_btn_NewChapter_inFront_bblk.OfficeImageId = "BevelShapeGallery"
        Me.grpRpt_mnu_btn_NewChapter_inFront_bblk.ShowImage = True
        '
        'grpRpt_mnu_btn_NewChapter_behind_bblk
        '
        Me.grpRpt_mnu_btn_NewChapter_behind_bblk.Label = "New Chapter (&Behind)"
        Me.grpRpt_mnu_btn_NewChapter_behind_bblk.Name = "grpRpt_mnu_btn_NewChapter_behind_bblk"
        Me.grpRpt_mnu_btn_NewChapter_behind_bblk.OfficeImageId = "BevelShapeGallery"
        Me.grpRpt_mnu_btn_NewChapter_behind_bblk.ShowImage = True
        Me.grpRpt_mnu_btn_NewChapter_behind_bblk.SuperTip = """Inserts a new chapter section behind the section that contains the current curso" &
    "r position."""
        '
        'grpRpt_mnu_Bibliography
        '
        Me.grpRpt_mnu_Bibliography.Items.Add(Me.grpOther_bibliography_bblk)
        Me.grpRpt_mnu_Bibliography.Items.Add(Me.grpOther_references_bblk)
        Me.grpRpt_mnu_Bibliography.Items.Add(Me.grpOther_worksCited_bblk)
        Me.grpRpt_mnu_Bibliography.KeyTip = "RB"
        Me.grpRpt_mnu_Bibliography.Label = "Bibliography"
        Me.grpRpt_mnu_Bibliography.Name = "grpRpt_mnu_Bibliography"
        Me.grpRpt_mnu_Bibliography.OfficeImageId = "BibliographyGallery"
        Me.grpRpt_mnu_Bibliography.ShowImage = True
        '
        'grpOther_bibliography_bblk
        '
        Me.grpOther_bibliography_bblk.Label = "&Bibliography"
        Me.grpOther_bibliography_bblk.Name = "grpOther_bibliography_bblk"
        Me.grpOther_bibliography_bblk.OfficeImageId = "ReplaceWithAutoText"
        Me.grpOther_bibliography_bblk.ShowImage = True
        '
        'grpOther_references_bblk
        '
        Me.grpOther_references_bblk.Label = "&References"
        Me.grpOther_references_bblk.Name = "grpOther_references_bblk"
        Me.grpOther_references_bblk.OfficeImageId = "ReplaceWithAutoText"
        Me.grpOther_references_bblk.ShowImage = True
        '
        'grpOther_worksCited_bblk
        '
        Me.grpOther_worksCited_bblk.Label = "&Works Cited"
        Me.grpOther_worksCited_bblk.Name = "grpOther_worksCited_bblk"
        Me.grpOther_worksCited_bblk.OfficeImageId = "ReplaceWithAutoText"
        Me.grpOther_worksCited_bblk.ShowImage = True
        '
        'grpReport_btn_ToggleView
        '
        Me.grpReport_btn_ToggleView.KeyTip = "RT"
        Me.grpReport_btn_ToggleView.Label = "Toggle View"
        Me.grpReport_btn_ToggleView.Name = "grpReport_btn_ToggleView"
        Me.grpReport_btn_ToggleView.OfficeImageId = "ContentControlBuildingBlockGallery"
        Me.grpReport_btn_ToggleView.ShowImage = True
        Me.grpReport_btn_ToggleView.SuperTip = """Turns all hidden characters and lines to visible/invisible. These characters and" &
    " lines do NOT print"""
        Me.grpReport_btn_ToggleView.Visible = False
        '
        'grpRpt_mnu_RefreshDocument
        '
        Me.grpRpt_mnu_RefreshDocument.Items.Add(Me.grpViewTools_Refresh_Stationery_Ref)
        Me.grpRpt_mnu_RefreshDocument.Items.Add(Me.Separator9)
        Me.grpRpt_mnu_RefreshDocument.Items.Add(Me.grpViewTools_Refresh_mnu_TOC)
        Me.grpRpt_mnu_RefreshDocument.Items.Add(Me.grpViewTools_Refresh_mnu_Chapters)
        Me.grpRpt_mnu_RefreshDocument.Items.Add(Me.grpViewTools_Refresh_mnu_Parts)
        Me.grpRpt_mnu_RefreshDocument.Items.Add(Me.Separator10)
        Me.grpRpt_mnu_RefreshDocument.Items.Add(Me.grpViewTools_Refresh_mnu_Tables)
        Me.grpRpt_mnu_RefreshDocument.Items.Add(Me.grpViewTools_Refresh_mnu_Figures)
        Me.grpRpt_mnu_RefreshDocument.Items.Add(Me.grpViewTools_Refresh_mnu_Boxes)
        Me.grpRpt_mnu_RefreshDocument.Items.Add(Me.Separator11)
        Me.grpRpt_mnu_RefreshDocument.Items.Add(Me.grpViewTools_Refresh_mnu_All)
        Me.grpRpt_mnu_RefreshDocument.Items.Add(Me.Separator12)
        Me.grpRpt_mnu_RefreshDocument.Items.Add(Me.grpViewTools_Refresh_mnu_Every)
        Me.grpRpt_mnu_RefreshDocument.Items.Add(Me.Separator13)
        Me.grpRpt_mnu_RefreshDocument.Items.Add(Me.mnu_grpViewTools_Refresh_btn_setRefFldNotBold)
        Me.grpRpt_mnu_RefreshDocument.KeyTip = "RD"
        Me.grpRpt_mnu_RefreshDocument.Label = "Refresh Document"
        Me.grpRpt_mnu_RefreshDocument.Name = "grpRpt_mnu_RefreshDocument"
        Me.grpRpt_mnu_RefreshDocument.OfficeImageId = "PictureBrightnessGallery"
        Me.grpRpt_mnu_RefreshDocument.ShowImage = True
        Me.grpRpt_mnu_RefreshDocument.SuperTip = """Will refresh selected fields in the document. Handy if the page numbers, TOC, Ta" &
    "bles, Figures or Boxes haven't automatically updated while you are working."""
        '
        'grpViewTools_Refresh_Stationery_Ref
        '
        Me.grpViewTools_Refresh_Stationery_Ref.Label = "&Stationery Reference Field (for letters, memos etc)"
        Me.grpViewTools_Refresh_Stationery_Ref.Name = "grpViewTools_Refresh_Stationery_Ref"
        Me.grpViewTools_Refresh_Stationery_Ref.OfficeImageId = "TableOfContentsGallery"
        Me.grpViewTools_Refresh_Stationery_Ref.ShowImage = True
        '
        'Separator9
        '
        Me.Separator9.Name = "Separator9"
        '
        'grpViewTools_Refresh_mnu_TOC
        '
        Me.grpViewTools_Refresh_mnu_TOC.Label = "&Table Of Contents"
        Me.grpViewTools_Refresh_mnu_TOC.Name = "grpViewTools_Refresh_mnu_TOC"
        Me.grpViewTools_Refresh_mnu_TOC.ShowImage = True
        '
        'grpViewTools_Refresh_mnu_Chapters
        '
        Me.grpViewTools_Refresh_mnu_Chapters.Label = "Chapter and Appendix &main headings"
        Me.grpViewTools_Refresh_mnu_Chapters.Name = "grpViewTools_Refresh_mnu_Chapters"
        Me.grpViewTools_Refresh_mnu_Chapters.ShowImage = True
        '
        'grpViewTools_Refresh_mnu_Parts
        '
        Me.grpViewTools_Refresh_mnu_Parts.Label = "&Part Numbers"
        Me.grpViewTools_Refresh_mnu_Parts.Name = "grpViewTools_Refresh_mnu_Parts"
        Me.grpViewTools_Refresh_mnu_Parts.ShowImage = True
        '
        'Separator10
        '
        Me.Separator10.Name = "Separator10"
        '
        'grpViewTools_Refresh_mnu_Tables
        '
        Me.grpViewTools_Refresh_mnu_Tables.Label = "&Table numbering"
        Me.grpViewTools_Refresh_mnu_Tables.Name = "grpViewTools_Refresh_mnu_Tables"
        Me.grpViewTools_Refresh_mnu_Tables.ShowImage = True
        '
        'grpViewTools_Refresh_mnu_Figures
        '
        Me.grpViewTools_Refresh_mnu_Figures.Label = "&Figure numbering"
        Me.grpViewTools_Refresh_mnu_Figures.Name = "grpViewTools_Refresh_mnu_Figures"
        Me.grpViewTools_Refresh_mnu_Figures.ShowImage = True
        '
        'grpViewTools_Refresh_mnu_Boxes
        '
        Me.grpViewTools_Refresh_mnu_Boxes.Label = "&Box numbering"
        Me.grpViewTools_Refresh_mnu_Boxes.Name = "grpViewTools_Refresh_mnu_Boxes"
        Me.grpViewTools_Refresh_mnu_Boxes.ShowImage = True
        '
        'Separator11
        '
        Me.Separator11.Name = "Separator11"
        '
        'grpViewTools_Refresh_mnu_All
        '
        Me.grpViewTools_Refresh_mnu_All.Label = "Do &All of theAbove"
        Me.grpViewTools_Refresh_mnu_All.Name = "grpViewTools_Refresh_mnu_All"
        Me.grpViewTools_Refresh_mnu_All.ShowImage = True
        '
        'Separator12
        '
        Me.Separator12.Name = "Separator12"
        '
        'grpViewTools_Refresh_mnu_Every
        '
        Me.grpViewTools_Refresh_mnu_Every.Label = "&Every Field in the Document"
        Me.grpViewTools_Refresh_mnu_Every.Name = "grpViewTools_Refresh_mnu_Every"
        Me.grpViewTools_Refresh_mnu_Every.ShowImage = True
        '
        'Separator13
        '
        Me.Separator13.Name = "Separator13"
        '
        'mnu_grpViewTools_Refresh_btn_setRefFldNotBold
        '
        Me.mnu_grpViewTools_Refresh_btn_setRefFldNotBold.Label = "Set all Cross Reference Fields to '&Not Bold'"
        Me.mnu_grpViewTools_Refresh_btn_setRefFldNotBold.Name = "mnu_grpViewTools_Refresh_btn_setRefFldNotBold"
        Me.mnu_grpViewTools_Refresh_btn_setRefFldNotBold.ShowImage = True
        '
        'grpRpt_mnu_ApplyColour
        '
        Me.grpRpt_mnu_ApplyColour.Items.Add(Me.grpReport_mnu01_SelectedText)
        Me.grpRpt_mnu_ApplyColour.Items.Add(Me.grpReport_mnu01_SelectedTblCells)
        Me.grpRpt_mnu_ApplyColour.Items.Add(Me.grpReport_mnu01_ImageBackPanel)
        Me.grpRpt_mnu_ApplyColour.Items.Add(Me.Separator7)
        Me.grpRpt_mnu_ApplyColour.Items.Add(Me.grpReport_mnu_CaseStudies_RecolourLogo_toWhite)
        Me.grpRpt_mnu_ApplyColour.Items.Add(Me.grpReport_mnu_CaseStudies_RecolourLogo_Reset)
        Me.grpRpt_mnu_ApplyColour.Items.Add(Me.Separator8)
        Me.grpRpt_mnu_ApplyColour.Items.Add(Me.grpReport_mnu_CaseStudies_RecolourFooter_toWhite)
        Me.grpRpt_mnu_ApplyColour.Items.Add(Me.grpReport_mnu_CaseStudies_RecolourFooter_Reset)
        Me.grpRpt_mnu_ApplyColour.KeyTip = "RL"
        Me.grpRpt_mnu_ApplyColour.Label = "Apply Colour"
        Me.grpRpt_mnu_ApplyColour.Name = "grpRpt_mnu_ApplyColour"
        Me.grpRpt_mnu_ApplyColour.OfficeImageId = "ViewBackToColorView"
        Me.grpRpt_mnu_ApplyColour.ShowImage = True
        Me.grpRpt_mnu_ApplyColour.SuperTip = """Apply a specific colour to the 'selected text', selected table cells, the image " &
    "back panel, AA logo or the footer text"""
        '
        'grpReport_mnu01_SelectedText
        '
        Me.grpReport_mnu01_SelectedText.Label = "Recolour &selected text"
        Me.grpReport_mnu01_SelectedText.Name = "grpReport_mnu01_SelectedText"
        Me.grpReport_mnu01_SelectedText.OfficeImageId = "BevelShapeGallery"
        Me.grpReport_mnu01_SelectedText.ShowImage = True
        '
        'grpReport_mnu01_SelectedTblCells
        '
        Me.grpReport_mnu01_SelectedTblCells.Label = "Recolour selected &table cells"
        Me.grpReport_mnu01_SelectedTblCells.Name = "grpReport_mnu01_SelectedTblCells"
        Me.grpReport_mnu01_SelectedTblCells.OfficeImageId = "BevelShapeGallery"
        Me.grpReport_mnu01_SelectedTblCells.ShowImage = True
        '
        'grpReport_mnu01_ImageBackPanel
        '
        Me.grpReport_mnu01_ImageBackPanel.Label = "Recolour image &back panel"
        Me.grpReport_mnu01_ImageBackPanel.Name = "grpReport_mnu01_ImageBackPanel"
        Me.grpReport_mnu01_ImageBackPanel.OfficeImageId = "BevelShapeGallery"
        Me.grpReport_mnu01_ImageBackPanel.ShowImage = True
        '
        'Separator7
        '
        Me.Separator7.Name = "Separator7"
        '
        'grpReport_mnu_CaseStudies_RecolourLogo_toWhite
        '
        Me.grpReport_mnu_CaseStudies_RecolourLogo_toWhite.Label = "Recolour logo to &white"
        Me.grpReport_mnu_CaseStudies_RecolourLogo_toWhite.Name = "grpReport_mnu_CaseStudies_RecolourLogo_toWhite"
        Me.grpReport_mnu_CaseStudies_RecolourLogo_toWhite.OfficeImageId = "TextBoxInsert"
        Me.grpReport_mnu_CaseStudies_RecolourLogo_toWhite.ShowImage = True
        '
        'grpReport_mnu_CaseStudies_RecolourLogo_Reset
        '
        Me.grpReport_mnu_CaseStudies_RecolourLogo_Reset.Label = "Reset the &logo colour"
        Me.grpReport_mnu_CaseStudies_RecolourLogo_Reset.Name = "grpReport_mnu_CaseStudies_RecolourLogo_Reset"
        Me.grpReport_mnu_CaseStudies_RecolourLogo_Reset.OfficeImageId = "TextBoxInsert"
        Me.grpReport_mnu_CaseStudies_RecolourLogo_Reset.ShowImage = True
        '
        'Separator8
        '
        Me.Separator8.Name = "Separator8"
        '
        'grpReport_mnu_CaseStudies_RecolourFooter_toWhite
        '
        Me.grpReport_mnu_CaseStudies_RecolourFooter_toWhite.Label = "Recolour &footer text to white"
        Me.grpReport_mnu_CaseStudies_RecolourFooter_toWhite.Name = "grpReport_mnu_CaseStudies_RecolourFooter_toWhite"
        Me.grpReport_mnu_CaseStudies_RecolourFooter_toWhite.OfficeImageId = "TextBoxInsert"
        Me.grpReport_mnu_CaseStudies_RecolourFooter_toWhite.ShowImage = True
        '
        'grpReport_mnu_CaseStudies_RecolourFooter_Reset
        '
        Me.grpReport_mnu_CaseStudies_RecolourFooter_Reset.Label = "&Reset footer text colour"
        Me.grpReport_mnu_CaseStudies_RecolourFooter_Reset.Name = "grpReport_mnu_CaseStudies_RecolourFooter_Reset"
        Me.grpReport_mnu_CaseStudies_RecolourFooter_Reset.OfficeImageId = "TextBoxInsert"
        Me.grpReport_mnu_CaseStudies_RecolourFooter_Reset.ShowImage = True
        '
        'grpRpt_Appendix
        '
        Me.grpRpt_Appendix.Items.Add(Me.grpAppendix_mnu01)
        Me.grpRpt_Appendix.Items.Add(Me.grpReport_mnu_NewAppAtt)
        Me.grpRpt_Appendix.Label = "Appendix"
        Me.grpRpt_Appendix.Name = "grpRpt_Appendix"
        '
        'grpAppendix_mnu01
        '
        Me.grpAppendix_mnu01.Image = Global.AA_GeneralReport_Addin.My.Resources.Resources.NewPart_TG
        Me.grpAppendix_mnu01.Items.Add(Me.grpAppendix_newAppPart)
        Me.grpAppendix_mnu01.Items.Add(Me.grpAppendix_newAttPart)
        Me.grpAppendix_mnu01.KeyTip = "AP"
        Me.grpAppendix_mnu01.Label = "Part Divider (App/Att)"
        Me.grpAppendix_mnu01.Name = "grpAppendix_mnu01"
        Me.grpAppendix_mnu01.ScreenTip = "Part Divider (App/Att)"
        Me.grpAppendix_mnu01.ShowImage = True
        Me.grpAppendix_mnu01.SuperTip = """You can select either an Appendix Divider or an Attachment Divider"""
        '
        'grpAppendix_newAppPart
        '
        Me.grpAppendix_newAppPart.Label = "&Appendix Divider"
        Me.grpAppendix_newAppPart.Name = "grpAppendix_newAppPart"
        Me.grpAppendix_newAppPart.OfficeImageId = "TextBoxInsert"
        Me.grpAppendix_newAppPart.ScreenTip = "Appendix Divider"
        Me.grpAppendix_newAppPart.ShowImage = True
        Me.grpAppendix_newAppPart.SuperTip = """Inserts a new Appendix Part Divider. All appendix chapters must be preceded by a" &
    " single Appendix Part Divider"""
        '
        'grpAppendix_newAttPart
        '
        Me.grpAppendix_newAttPart.Label = "A&ttachment Divider"
        Me.grpAppendix_newAttPart.Name = "grpAppendix_newAttPart"
        Me.grpAppendix_newAttPart.OfficeImageId = "TextBoxInsert"
        Me.grpAppendix_newAttPart.ScreenTip = "Attachment Divider"
        Me.grpAppendix_newAttPart.ShowImage = True
        Me.grpAppendix_newAttPart.SuperTip = """Inserts a new Attachment Part Divider. All attachment chapters must be preced by" &
    " a single Attachment Part Divider"""
        '
        'grpReport_mnu_NewAppAtt
        '
        Me.grpReport_mnu_NewAppAtt.Items.Add(Me.grpAppendix_newAppChapter_inFront_bblk)
        Me.grpReport_mnu_NewAppAtt.Items.Add(Me.grpAppendix_newAppChapter_behind_bblk)
        Me.grpReport_mnu_NewAppAtt.KeyTip = "AC"
        Me.grpReport_mnu_NewAppAtt.Label = "New Appendix/Att"
        Me.grpReport_mnu_NewAppAtt.Name = "grpReport_mnu_NewAppAtt"
        Me.grpReport_mnu_NewAppAtt.OfficeImageId = "TextBoxInsert"
        Me.grpReport_mnu_NewAppAtt.ScreenTip = "New Appendix/Att"
        Me.grpReport_mnu_NewAppAtt.ShowImage = True
        Me.grpReport_mnu_NewAppAtt.SuperTip = " ""Inserts a new page with Appendix/Attachment heading separated by a section brea" &
    "k. Page number and field in footer are linked to Appendix heading."""
        '
        'grpAppendix_newAppChapter_inFront_bblk
        '
        Me.grpAppendix_newAppChapter_inFront_bblk.Label = "New App/Att (in &Front)"
        Me.grpAppendix_newAppChapter_inFront_bblk.Name = "grpAppendix_newAppChapter_inFront_bblk"
        Me.grpAppendix_newAppChapter_inFront_bblk.OfficeImageId = "TextBoxInsert"
        Me.grpAppendix_newAppChapter_inFront_bblk.ScreenTip = "New App/Att (in Front)"
        Me.grpAppendix_newAppChapter_inFront_bblk.ShowImage = True
        Me.grpAppendix_newAppChapter_inFront_bblk.SuperTip = """Inserts a new App/Att section in front of the Section that contains the current " &
    "cursor position."""
        '
        'grpAppendix_newAppChapter_behind_bblk
        '
        Me.grpAppendix_newAppChapter_behind_bblk.Label = "New App/Att (&Behind)"
        Me.grpAppendix_newAppChapter_behind_bblk.Name = "grpAppendix_newAppChapter_behind_bblk"
        Me.grpAppendix_newAppChapter_behind_bblk.OfficeImageId = "TextBoxInsert"
        Me.grpAppendix_newAppChapter_behind_bblk.ScreenTip = "New App/Att (Behind)"
        Me.grpAppendix_newAppChapter_behind_bblk.ShowImage = True
        Me.grpAppendix_newAppChapter_behind_bblk.SuperTip = """Inserts a new App/Attsection behind the Section that contains the current cursor" &
    " position."""
        '
        'grpRpt_sectOptions
        '
        Me.grpRpt_sectOptions.Items.Add(Me.mnuCloseDocuments000)
        Me.grpRpt_sectOptions.Items.Add(Me.grpRpt_sectOptions_btn_delSection)
        Me.grpRpt_sectOptions.Items.Add(Me.grpOther_mnuHFS)
        Me.grpRpt_sectOptions.Items.Add(Me.grpSectOptions_mnu_ResetLndPrt)
        Me.grpRpt_sectOptions.Items.Add(Me.grpSectOptions_mnu_ResetResizeLandscape)
        Me.grpRpt_sectOptions.Items.Add(Me.mnu_grpReport_Columns)
        Me.grpRpt_sectOptions.Label = "Other Section Options"
        Me.grpRpt_sectOptions.Name = "grpRpt_sectOptions"
        '
        'mnuCloseDocuments000
        '
        Me.mnuCloseDocuments000.Items.Add(Me.grpSectOptions_submnu_LndWidthOptions)
        Me.mnuCloseDocuments000.Items.Add(Me.Menu3)
        Me.mnuCloseDocuments000.Items.Add(Me.Separator14)
        Me.mnuCloseDocuments000.Items.Add(Me.grpSectOptions_sect_InsertSection_AtSelection)
        Me.mnuCloseDocuments000.KeyTip = "OS"
        Me.mnuCloseDocuments000.Label = "Insert blank Lnd/Prt Section"
        Me.mnuCloseDocuments000.Name = "mnuCloseDocuments000"
        Me.mnuCloseDocuments000.OfficeImageId = "BlackAndWhiteWhite"
        Me.mnuCloseDocuments000.ScreenTip = "Insert blank Lnd/Prt Section"
        Me.mnuCloseDocuments000.ShowImage = True
        Me.mnuCloseDocuments000.SuperTip = """Use a menu item to insert a blank section either at the cursor location or at th" &
    "e end of the document."""
        '
        'grpSectOptions_submnu_LndWidthOptions
        '
        Me.grpSectOptions_submnu_LndWidthOptions.Items.Add(Me.grpSectOptions_sect_InsertSectionBounded_Lnd)
        Me.grpSectOptions_submnu_LndWidthOptions.Items.Add(Me.grpSectOptions_sect_InsertSectionBounded_Lnd_wide)
        Me.grpSectOptions_submnu_LndWidthOptions.Label = "Insert Blank 'Bounded' Section (&Landscape)"
        Me.grpSectOptions_submnu_LndWidthOptions.Name = "grpSectOptions_submnu_LndWidthOptions"
        Me.grpSectOptions_submnu_LndWidthOptions.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_submnu_LndWidthOptions.ShowImage = True
        '
        'grpSectOptions_sect_InsertSectionBounded_Lnd
        '
        Me.grpSectOptions_sect_InsertSectionBounded_Lnd.Label = "&Standard margins"
        Me.grpSectOptions_sect_InsertSectionBounded_Lnd.Name = "grpSectOptions_sect_InsertSectionBounded_Lnd"
        Me.grpSectOptions_sect_InsertSectionBounded_Lnd.OfficeImageId = "FileNew"
        Me.grpSectOptions_sect_InsertSectionBounded_Lnd.ShowImage = True
        '
        'grpSectOptions_sect_InsertSectionBounded_Lnd_wide
        '
        Me.grpSectOptions_sect_InsertSectionBounded_Lnd_wide.Label = "&Wide margins"
        Me.grpSectOptions_sect_InsertSectionBounded_Lnd_wide.Name = "grpSectOptions_sect_InsertSectionBounded_Lnd_wide"
        Me.grpSectOptions_sect_InsertSectionBounded_Lnd_wide.OfficeImageId = "FileNew"
        Me.grpSectOptions_sect_InsertSectionBounded_Lnd_wide.ShowImage = True
        '
        'Menu3
        '
        Me.Menu3.Items.Add(Me.grpSectOptions_sect_InsertSectionBounded_Prt)
        Me.Menu3.Items.Add(Me.grpSectOptions_sect_InsertSectionBounded_Prt_wide)
        Me.Menu3.Label = "Insert Blank 'Bounded' Section (&Portrait)"
        Me.Menu3.Name = "Menu3"
        Me.Menu3.OfficeImageId = "BevelShapeGallery"
        Me.Menu3.ShowImage = True
        '
        'grpSectOptions_sect_InsertSectionBounded_Prt
        '
        Me.grpSectOptions_sect_InsertSectionBounded_Prt.Label = "&Standard margins"
        Me.grpSectOptions_sect_InsertSectionBounded_Prt.Name = "grpSectOptions_sect_InsertSectionBounded_Prt"
        Me.grpSectOptions_sect_InsertSectionBounded_Prt.OfficeImageId = "FileNew"
        Me.grpSectOptions_sect_InsertSectionBounded_Prt.ShowImage = True
        Me.grpSectOptions_sect_InsertSectionBounded_Prt.SuperTip = resources.GetString("grpSectOptions_sect_InsertSectionBounded_Prt.SuperTip")
        '
        'grpSectOptions_sect_InsertSectionBounded_Prt_wide
        '
        Me.grpSectOptions_sect_InsertSectionBounded_Prt_wide.Label = "&Wide margins"
        Me.grpSectOptions_sect_InsertSectionBounded_Prt_wide.Name = "grpSectOptions_sect_InsertSectionBounded_Prt_wide"
        Me.grpSectOptions_sect_InsertSectionBounded_Prt_wide.OfficeImageId = "FileNew"
        Me.grpSectOptions_sect_InsertSectionBounded_Prt_wide.ShowImage = True
        '
        'Separator14
        '
        Me.Separator14.Name = "Separator14"
        '
        'grpSectOptions_sect_InsertSection_AtSelection
        '
        Me.grpSectOptions_sect_InsertSection_AtSelection.Label = "Insert Blank Section at &Selection"
        Me.grpSectOptions_sect_InsertSection_AtSelection.Name = "grpSectOptions_sect_InsertSection_AtSelection"
        Me.grpSectOptions_sect_InsertSection_AtSelection.OfficeImageId = "FileNew"
        Me.grpSectOptions_sect_InsertSection_AtSelection.ScreenTip = "Insert Blank Section at &Selection"
        Me.grpSectOptions_sect_InsertSection_AtSelection.ShowImage = True
        Me.grpSectOptions_sect_InsertSection_AtSelection.SuperTip = """This menu item inserts a 'section break' at the current cursor position. When fi" &
    "nished, the selection is at the beginning of the new section."""
        '
        'grpRpt_sectOptions_btn_delSection
        '
        Me.grpRpt_sectOptions_btn_delSection.KeyTip = "OX"
        Me.grpRpt_sectOptions_btn_delSection.Label = "DELETE section"
        Me.grpRpt_sectOptions_btn_delSection.Name = "grpRpt_sectOptions_btn_delSection"
        Me.grpRpt_sectOptions_btn_delSection.OfficeImageId = "DeleteWeb"
        Me.grpRpt_sectOptions_btn_delSection.ScreenTip = "DELETE section"
        Me.grpRpt_sectOptions_btn_delSection.ShowImage = True
        Me.grpRpt_sectOptions_btn_delSection.SuperTip = """This button deletes the section that contains your current cursor. Note that suc" &
    "cessive application of the Word 'undo' function will generally retrieve the dele" &
    "ted section """
        '
        'grpOther_mnuHFS
        '
        Me.grpOther_mnuHFS.Items.Add(Me.grpSectOptions_header_ClearTextandShapes)
        Me.grpOther_mnuHFS.Items.Add(Me.grpSectOptions_footer_ClearText)
        Me.grpOther_mnuHFS.Items.Add(Me.grpSectOptions_footer_ClearTextandPageNum)
        Me.grpOther_mnuHFS.Items.Add(Me.Separator15)
        Me.grpOther_mnuHFS.Items.Add(Me.grpSectOptions_footer_clearSubTitleField)
        Me.grpOther_mnuHFS.Items.Add(Me.Separator16)
        Me.grpOther_mnuHFS.Items.Add(Me.grpSectOptions_footer_resetText)
        Me.grpOther_mnuHFS.Items.Add(Me.Separator17)
        Me.grpOther_mnuHFS.Items.Add(Me.grpOther_mnuHFS_sub00_restoreHF)
        Me.grpOther_mnuHFS.KeyTip = "OH"
        Me.grpOther_mnuHFS.Label = "Header/Footers"
        Me.grpOther_mnuHFS.Name = "grpOther_mnuHFS"
        Me.grpOther_mnuHFS.OfficeImageId = "BevelShapeGallery"
        Me.grpOther_mnuHFS.ScreenTip = "Header/Footers"
        Me.grpOther_mnuHFS.ShowImage = True
        Me.grpOther_mnuHFS.SuperTip = """Use these menu items to clear or restore the headers and/or the footers of the s" &
    "elected section."""
        '
        'grpSectOptions_header_ClearTextandShapes
        '
        Me.grpSectOptions_header_ClearTextandShapes.Label = "Clear current &header (Text and Shapes)"
        Me.grpSectOptions_header_ClearTextandShapes.Name = "grpSectOptions_header_ClearTextandShapes"
        Me.grpSectOptions_header_ClearTextandShapes.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_header_ClearTextandShapes.ScreenTip = "Clear current header (Text and Shapes)"
        Me.grpSectOptions_header_ClearTextandShapes.ShowImage = True
        Me.grpSectOptions_header_ClearTextandShapes.SuperTip = """This button will remove any text and branding in the Header of the current secti" &
    "on."""
        '
        'grpSectOptions_footer_ClearText
        '
        Me.grpSectOptions_footer_ClearText.Label = "Clear current &footer (Text)"
        Me.grpSectOptions_footer_ClearText.Name = "grpSectOptions_footer_ClearText"
        Me.grpSectOptions_footer_ClearText.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_footer_ClearText.ScreenTip = "Clear current footer (Text)"
        Me.grpSectOptions_footer_ClearText.ShowImage = True
        Me.grpSectOptions_footer_ClearText.SuperTip = """This button will remove any text in the footer of the current section. The page " &
    "number will be left as is """
        '
        'grpSectOptions_footer_ClearTextandPageNum
        '
        Me.grpSectOptions_footer_ClearTextandPageNum.Label = """Clear current footer (Text and &Page #)"""
        Me.grpSectOptions_footer_ClearTextandPageNum.Name = "grpSectOptions_footer_ClearTextandPageNum"
        Me.grpSectOptions_footer_ClearTextandPageNum.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_footer_ClearTextandPageNum.ScreenTip = """Clear current footer (Text and Page #)"""
        Me.grpSectOptions_footer_ClearTextandPageNum.ShowImage = True
        Me.grpSectOptions_footer_ClearTextandPageNum.SuperTip = """This button will remove any text in the footer of the current section. The page " &
    "number will also be removed """
        '
        'Separator15
        '
        Me.Separator15.Name = "Separator15"
        '
        'grpSectOptions_footer_clearSubTitleField
        '
        Me.grpSectOptions_footer_clearSubTitleField.Label = "Clear &all footer(s) of the Sub Title Reference"
        Me.grpSectOptions_footer_clearSubTitleField.Name = "grpSectOptions_footer_clearSubTitleField"
        Me.grpSectOptions_footer_clearSubTitleField.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_footer_clearSubTitleField.ScreenTip = "Clear all footer(s) of the Sub Title Reference"
        Me.grpSectOptions_footer_clearSubTitleField.ShowImage = True
        Me.grpSectOptions_footer_clearSubTitleField.SuperTip = """This button will remove any text in the footers that references back to the Sub " &
    "Title text in the Cover Page."""
        '
        'Separator16
        '
        Me.Separator16.Name = "Separator16"
        '
        'grpSectOptions_footer_resetText
        '
        Me.grpSectOptions_footer_resetText.Label = "Reset &all footer(s) text to the new two line standard"
        Me.grpSectOptions_footer_resetText.Name = "grpSectOptions_footer_resetText"
        Me.grpSectOptions_footer_resetText.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_footer_resetText.ScreenTip = "Reset all footer(s) text to the new two line standard"
        Me.grpSectOptions_footer_resetText.ShowImage = True
        Me.grpSectOptions_footer_resetText.SuperTip = """This button will remove any text in the footers of the current document and then" &
    " will reset that text to the current ACIL Allen standard."""
        '
        'Separator17
        '
        Me.Separator17.Name = "Separator17"
        '
        'grpOther_mnuHFS_sub00_restoreHF
        '
        Me.grpOther_mnuHFS_sub00_restoreHF.Items.Add(Me.grpSectOptions_hfs_restoreHF_ES)
        Me.grpOther_mnuHFS_sub00_restoreHF.Items.Add(Me.grpSectOptions_hfs_restoreHF_RP)
        Me.grpOther_mnuHFS_sub00_restoreHF.Items.Add(Me.grpSectOptions_hfs_restoreHF_AP)
        Me.grpOther_mnuHFS_sub00_restoreHF.Label = "&Restore Headers and Footers"
        Me.grpOther_mnuHFS_sub00_restoreHF.Name = "grpOther_mnuHFS_sub00_restoreHF"
        Me.grpOther_mnuHFS_sub00_restoreHF.OfficeImageId = "BevelShapeGallery"
        Me.grpOther_mnuHFS_sub00_restoreHF.ScreenTip = "Restore Headers and Footers"
        Me.grpOther_mnuHFS_sub00_restoreHF.ShowImage = True
        Me.grpOther_mnuHFS_sub00_restoreHF.SuperTip = """Use these menu items to restore the headers and footers of the current section t" &
    "o the specified format"""
        '
        'grpSectOptions_hfs_restoreHF_ES
        '
        Me.grpSectOptions_hfs_restoreHF_ES.Label = "Restore &Executive Summary (ES) and Glossary HF"
        Me.grpSectOptions_hfs_restoreHF_ES.Name = "grpSectOptions_hfs_restoreHF_ES"
        Me.grpSectOptions_hfs_restoreHF_ES.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_hfs_restoreHF_ES.ScreenTip = "Restore Executive Summary (ES) and Glossary HF"
        Me.grpSectOptions_hfs_restoreHF_ES.ShowImage = True
        Me.grpSectOptions_hfs_restoreHF_ES.SuperTip = """This button restores the headers and footers of the current section to the Acil " &
    "Allen Executive Summary (ES) / Glossary standard """
        '
        'grpSectOptions_hfs_restoreHF_RP
        '
        Me.grpSectOptions_hfs_restoreHF_RP.Label = "Restore &Report/Brief Body (RP/Brief) HF"
        Me.grpSectOptions_hfs_restoreHF_RP.Name = "grpSectOptions_hfs_restoreHF_RP"
        Me.grpSectOptions_hfs_restoreHF_RP.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_hfs_restoreHF_RP.ScreenTip = "Restore Report/Brief Body (RP/Brief) HF"
        Me.grpSectOptions_hfs_restoreHF_RP.ShowImage = True
        Me.grpSectOptions_hfs_restoreHF_RP.SuperTip = """This button restores the headers and footers of the current section to the Acil " &
    "Allen Report/Brief Body (RP/Brief) standard """
        '
        'grpSectOptions_hfs_restoreHF_AP
        '
        Me.grpSectOptions_hfs_restoreHF_AP.Label = "Restore &Appendix (AP) HF"
        Me.grpSectOptions_hfs_restoreHF_AP.Name = "grpSectOptions_hfs_restoreHF_AP"
        Me.grpSectOptions_hfs_restoreHF_AP.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_hfs_restoreHF_AP.ScreenTip = "Restore Appendix (AP) HF"
        Me.grpSectOptions_hfs_restoreHF_AP.ShowImage = True
        Me.grpSectOptions_hfs_restoreHF_AP.SuperTip = """This button restores the headers and footers of the current section to the Acil " &
    "Allen Appendix (AP) standard """
        '
        'grpSectOptions_mnu_ResetLndPrt
        '
        Me.grpSectOptions_mnu_ResetLndPrt.Items.Add(Me.grpSectOptions_mnu_sub00_ResetLndPrt_Lnd)
        Me.grpSectOptions_mnu_ResetLndPrt.Items.Add(Me.grpSectOptions_mnu_sub01_ResetLndPrt_Ln)
        Me.grpSectOptions_mnu_ResetLndPrt.KeyTip = "OE"
        Me.grpSectOptions_mnu_ResetLndPrt.Label = "Reset to Lnd/Prt defaults"
        Me.grpSectOptions_mnu_ResetLndPrt.Name = "grpSectOptions_mnu_ResetLndPrt"
        Me.grpSectOptions_mnu_ResetLndPrt.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_mnu_ResetLndPrt.ScreenTip = "Reset to Lnd/Prt defaults"
        Me.grpSectOptions_mnu_ResetLndPrt.ShowImage = True
        Me.grpSectOptions_mnu_ResetLndPrt.SuperTip = """'Reset' will adjust the 'page' measurements back to the standard values and inse" &
    "rt new 'Header/Footers' """
        '
        'grpSectOptions_mnu_sub00_ResetLndPrt_Lnd
        '
        Me.grpSectOptions_mnu_sub00_ResetLndPrt_Lnd.Items.Add(Me.grpSectOptions_resetTo_Lnd_ES)
        Me.grpSectOptions_mnu_sub00_ResetLndPrt_Lnd.Items.Add(Me.grpSectOptions_resetTo_Lnd_RP)
        Me.grpSectOptions_mnu_sub00_ResetLndPrt_Lnd.Items.Add(Me.grpSectOptions_resetTo_Lnd_AP)
        Me.grpSectOptions_mnu_sub00_ResetLndPrt_Lnd.Label = "Reset to &Landscape"
        Me.grpSectOptions_mnu_sub00_ResetLndPrt_Lnd.Name = "grpSectOptions_mnu_sub00_ResetLndPrt_Lnd"
        Me.grpSectOptions_mnu_sub00_ResetLndPrt_Lnd.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_mnu_sub00_ResetLndPrt_Lnd.ScreenTip = "Reset to Landscape"
        Me.grpSectOptions_mnu_sub00_ResetLndPrt_Lnd.ShowImage = True
        Me.grpSectOptions_mnu_sub00_ResetLndPrt_Lnd.SuperTip = """Use this menu item to restore the formatting of the current section. Choose Exec" &
    "utive Summary, Report Body or Appendix depending on where in the report your cur" &
    "sor is located."""
        '
        'grpSectOptions_resetTo_Lnd_ES
        '
        Me.grpSectOptions_resetTo_Lnd_ES.Label = "Reset to &Executive Summary (ES) Landscape default"
        Me.grpSectOptions_resetTo_Lnd_ES.Name = "grpSectOptions_resetTo_Lnd_ES"
        Me.grpSectOptions_resetTo_Lnd_ES.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_resetTo_Lnd_ES.ScreenTip = "Reset to Executive Summary (ES) Landscape default"
        Me.grpSectOptions_resetTo_Lnd_ES.ShowImage = True
        Me.grpSectOptions_resetTo_Lnd_ES.SuperTip = """This button will reset the current section to the default Landscape (ES) format," &
    " replacing the headers and footers"""
        '
        'grpSectOptions_resetTo_Lnd_RP
        '
        Me.grpSectOptions_resetTo_Lnd_RP.Label = "Reset to &Report/Brief Body (RP/Brief) Landscape default"
        Me.grpSectOptions_resetTo_Lnd_RP.Name = "grpSectOptions_resetTo_Lnd_RP"
        Me.grpSectOptions_resetTo_Lnd_RP.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_resetTo_Lnd_RP.ScreenTip = "Reset to Report/Brief Body (RP/Brief) Landscape default"
        Me.grpSectOptions_resetTo_Lnd_RP.ShowImage = True
        Me.grpSectOptions_resetTo_Lnd_RP.SuperTip = """This button will reset the current section to the default Landscape (RP/Brief) f" &
    "ormat, replacing the headers and footers"""
        '
        'grpSectOptions_resetTo_Lnd_AP
        '
        Me.grpSectOptions_resetTo_Lnd_AP.Label = "Reset to &Appendix (AP) Landscape default"
        Me.grpSectOptions_resetTo_Lnd_AP.Name = "grpSectOptions_resetTo_Lnd_AP"
        Me.grpSectOptions_resetTo_Lnd_AP.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_resetTo_Lnd_AP.ScreenTip = "Reset to Appendix (AP) Landscape default"
        Me.grpSectOptions_resetTo_Lnd_AP.ShowImage = True
        Me.grpSectOptions_resetTo_Lnd_AP.SuperTip = """This button will reset the current section to the default Landscape (AP) format," &
    " replacing the headers and footers"""
        '
        'grpSectOptions_mnu_sub01_ResetLndPrt_Ln
        '
        Me.grpSectOptions_mnu_sub01_ResetLndPrt_Ln.Items.Add(Me.grpSectOptions_resetTo_Prt_ES)
        Me.grpSectOptions_mnu_sub01_ResetLndPrt_Ln.Items.Add(Me.grpSectOptions_resetTo_Prt_RP)
        Me.grpSectOptions_mnu_sub01_ResetLndPrt_Ln.Items.Add(Me.grpSectOptions_resetTo_Prt_AP)
        Me.grpSectOptions_mnu_sub01_ResetLndPrt_Ln.Label = "Reset to &Portrait"
        Me.grpSectOptions_mnu_sub01_ResetLndPrt_Ln.Name = "grpSectOptions_mnu_sub01_ResetLndPrt_Ln"
        Me.grpSectOptions_mnu_sub01_ResetLndPrt_Ln.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_mnu_sub01_ResetLndPrt_Ln.ShowImage = True
        Me.grpSectOptions_mnu_sub01_ResetLndPrt_Ln.SuperTip = """Use this menu item to restore the formatting of the current section. Choose Exec" &
    "utive Summary, Report Body or Appendix depending on where in the report your cur" &
    "sor is located."""
        '
        'grpSectOptions_resetTo_Prt_ES
        '
        Me.grpSectOptions_resetTo_Prt_ES.Label = "Reset to Executive Summary (ES) Portrait default"
        Me.grpSectOptions_resetTo_Prt_ES.Name = "grpSectOptions_resetTo_Prt_ES"
        Me.grpSectOptions_resetTo_Prt_ES.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_resetTo_Prt_ES.ScreenTip = "Reset to Executive Summary (ES) Portrait default"
        Me.grpSectOptions_resetTo_Prt_ES.ShowImage = True
        Me.grpSectOptions_resetTo_Prt_ES.SuperTip = """This button will reset the current section to the default Portrait (ES) format, " &
    "replacing the headers and footers"""
        '
        'grpSectOptions_resetTo_Prt_RP
        '
        Me.grpSectOptions_resetTo_Prt_RP.Label = "Reset to &Report/Brief Body (RP/Brief) Portrait default"
        Me.grpSectOptions_resetTo_Prt_RP.Name = "grpSectOptions_resetTo_Prt_RP"
        Me.grpSectOptions_resetTo_Prt_RP.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_resetTo_Prt_RP.ScreenTip = "Reset to Report/Brief Body (RP/Brief) Portrait default"
        Me.grpSectOptions_resetTo_Prt_RP.ShowImage = True
        Me.grpSectOptions_resetTo_Prt_RP.SuperTip = """This button will reset the current section to the default Portrait (RP/Brief) fo" &
    "rmat, replacing the headers and footers"""
        '
        'grpSectOptions_resetTo_Prt_AP
        '
        Me.grpSectOptions_resetTo_Prt_AP.Label = "Reset to &Appendix (AP) Portrait default"
        Me.grpSectOptions_resetTo_Prt_AP.Name = "grpSectOptions_resetTo_Prt_AP"
        Me.grpSectOptions_resetTo_Prt_AP.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_resetTo_Prt_AP.ScreenTip = "Reset to Appendix (AP) Portrait default"
        Me.grpSectOptions_resetTo_Prt_AP.ShowImage = True
        Me.grpSectOptions_resetTo_Prt_AP.SuperTip = """This button will reset the current section to the default Portrait (AP) format, " &
    "replacing the headers and footers"""
        '
        'grpSectOptions_mnu_ResetResizeLandscape
        '
        Me.grpSectOptions_mnu_ResetResizeLandscape.Items.Add(Me.grpSectOptions_resizeTo_Landscape)
        Me.grpSectOptions_mnu_ResetResizeLandscape.Items.Add(Me.grpSectOptions_resizeTo_Portrait)
        Me.grpSectOptions_mnu_ResetResizeLandscape.Items.Add(Me.Separator18)
        Me.grpSectOptions_mnu_ResetResizeLandscape.Items.Add(Me.grpSectOptions_resize_toggleWidth)
        Me.grpSectOptions_mnu_ResetResizeLandscape.KeyTip = "OR"
        Me.grpSectOptions_mnu_ResetResizeLandscape.Label = "Re-Orient to Lnd/Prt"
        Me.grpSectOptions_mnu_ResetResizeLandscape.Name = "grpSectOptions_mnu_ResetResizeLandscape"
        Me.grpSectOptions_mnu_ResetResizeLandscape.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_mnu_ResetResizeLandscape.ScreenTip = "Re-Orient to Lnd/Prt"
        Me.grpSectOptions_mnu_ResetResizeLandscape.ShowImage = True
        Me.grpSectOptions_mnu_ResetResizeLandscape.SuperTip = """Use these menu items to re-orient the current section. The existing 'Header/Foot" &
    "ers' are kept intact and just resized"""
        '
        'grpSectOptions_resizeTo_Landscape
        '
        Me.grpSectOptions_resizeTo_Landscape.Label = "Re-Orient to &Landscape"
        Me.grpSectOptions_resizeTo_Landscape.Name = "grpSectOptions_resizeTo_Landscape"
        Me.grpSectOptions_resizeTo_Landscape.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_resizeTo_Landscape.ScreenTip = "Re-Orient to &Landscape"
        Me.grpSectOptions_resizeTo_Landscape.ShowImage = True
        Me.grpSectOptions_resizeTo_Landscape.SuperTip = """This button will re-orient the current section to Landscape format and applies t" &
    "he narrow (default) margins option."""
        '
        'grpSectOptions_resizeTo_Portrait
        '
        Me.grpSectOptions_resizeTo_Portrait.Label = "Re-Orient to &Portrait"
        Me.grpSectOptions_resizeTo_Portrait.Name = "grpSectOptions_resizeTo_Portrait"
        Me.grpSectOptions_resizeTo_Portrait.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_resizeTo_Portrait.ScreenTip = "Re-Orient to Portrait"
        Me.grpSectOptions_resizeTo_Portrait.ShowImage = True
        Me.grpSectOptions_resizeTo_Portrait.SuperTip = """This button will re-orient the current section to Portrait format and applies th" &
    "e narrow (default) margins option."""
        '
        'Separator18
        '
        Me.Separator18.Name = "Separator18"
        '
        'grpSectOptions_resize_toggleWidth
        '
        Me.grpSectOptions_resize_toggleWidth.Label = "&Toggle Margins"
        Me.grpSectOptions_resize_toggleWidth.Name = "grpSectOptions_resize_toggleWidth"
        Me.grpSectOptions_resize_toggleWidth.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_resize_toggleWidth.ScreenTip = "Toggle Margins"
        Me.grpSectOptions_resize_toggleWidth.ShowImage = True
        Me.grpSectOptions_resize_toggleWidth.SuperTip = """This button will toggle the margins of the current section between Wide and Stan" &
    "dard."""
        '
        'mnu_grpReport_Columns
        '
        Me.mnu_grpReport_Columns.Items.Add(Me.grpReport_Columns_04)
        Me.mnu_grpReport_Columns.Items.Add(Me.grpReport_Columns_03)
        Me.mnu_grpReport_Columns.Items.Add(Me.grpReport_Columns_02)
        Me.mnu_grpReport_Columns.Items.Add(Me.grpReport_Columns_02_LeftWide)
        Me.mnu_grpReport_Columns.Items.Add(Me.grpReport_Columns_02_RightWide)
        Me.mnu_grpReport_Columns.Items.Add(Me.Separator19)
        Me.mnu_grpReport_Columns.Items.Add(Me.grpReport_Columns_01)
        Me.mnu_grpReport_Columns.KeyTip = "OC"
        Me.mnu_grpReport_Columns.Label = "Columns Setup"
        Me.mnu_grpReport_Columns.Name = "mnu_grpReport_Columns"
        Me.mnu_grpReport_Columns.OfficeImageId = "ColumnsDialog"
        Me.mnu_grpReport_Columns.ScreenTip = "Columns Setup"
        Me.mnu_grpReport_Columns.ShowImage = True
        '
        'grpReport_Columns_04
        '
        Me.grpReport_Columns_04.Image = Global.AA_GeneralReport_Addin.My.Resources.Resources.icons_columns_4
        Me.grpReport_Columns_04.Label = "&4 Columns"
        Me.grpReport_Columns_04.Name = "grpReport_Columns_04"
        Me.grpReport_Columns_04.ScreenTip = "4 Columns"
        Me.grpReport_Columns_04.ShowImage = True
        Me.grpReport_Columns_04.SuperTip = """The current section will be setup with four equally sized columns"""
        '
        'grpReport_Columns_03
        '
        Me.grpReport_Columns_03.Image = Global.AA_GeneralReport_Addin.My.Resources.Resources.icons_columns_3
        Me.grpReport_Columns_03.Label = "&3 Columns"
        Me.grpReport_Columns_03.Name = "grpReport_Columns_03"
        Me.grpReport_Columns_03.ScreenTip = "3 Columns"
        Me.grpReport_Columns_03.ShowImage = True
        Me.grpReport_Columns_03.SuperTip = """The current section will be setup with three equally sized columns"""
        '
        'grpReport_Columns_02
        '
        Me.grpReport_Columns_02.Image = Global.AA_GeneralReport_Addin.My.Resources.Resources.icons_columns_2
        Me.grpReport_Columns_02.Label = "& 2 Columns"
        Me.grpReport_Columns_02.Name = "grpReport_Columns_02"
        Me.grpReport_Columns_02.OfficeImageId = "BevelShapeGallery"
        Me.grpReport_Columns_02.ScreenTip = "2 Columns"
        Me.grpReport_Columns_02.ShowImage = True
        Me.grpReport_Columns_02.SuperTip = """The current section will be setup with two equally sized columns"""
        '
        'grpReport_Columns_02_LeftWide
        '
        Me.grpReport_Columns_02_LeftWide.Image = Global.AA_GeneralReport_Addin.My.Resources.Resources.icons_columns_2_left
        Me.grpReport_Columns_02_LeftWide.Label = "2 Columns (&Left)"
        Me.grpReport_Columns_02_LeftWide.Name = "grpReport_Columns_02_LeftWide"
        Me.grpReport_Columns_02_LeftWide.ScreenTip = "2 Columns (Left)"
        Me.grpReport_Columns_02_LeftWide.ShowImage = True
        Me.grpReport_Columns_02_LeftWide.SuperTip = """The current section will be setup with two columns with the left column wider th" &
    "an the right"""
        '
        'grpReport_Columns_02_RightWide
        '
        Me.grpReport_Columns_02_RightWide.Image = Global.AA_GeneralReport_Addin.My.Resources.Resources.icons_columns_2_right
        Me.grpReport_Columns_02_RightWide.Label = "2 Columns (&Right)"
        Me.grpReport_Columns_02_RightWide.Name = "grpReport_Columns_02_RightWide"
        Me.grpReport_Columns_02_RightWide.ScreenTip = "2 Columns (Right)"
        Me.grpReport_Columns_02_RightWide.ShowImage = True
        Me.grpReport_Columns_02_RightWide.SuperTip = """The current section will be setup with two columns with the right column wider t" &
    "han the left"""
        '
        'Separator19
        '
        Me.Separator19.Name = "Separator19"
        '
        'grpReport_Columns_01
        '
        Me.grpReport_Columns_01.Label = "&1 Column"
        Me.grpReport_Columns_01.Name = "grpReport_Columns_01"
        Me.grpReport_Columns_01.OfficeImageId = "AlignJustifyMenu"
        Me.grpReport_Columns_01.ScreenTip = "1 Column"
        Me.grpReport_Columns_01.ShowImage = True
        Me.grpReport_Columns_01.SuperTip = """The current section will be setup with one column (i.e. a standard page)"""
        '
        'grpRpt_CoveringLetter
        '
        Me.grpRpt_CoveringLetter.Items.Add(Me.grpLetter_insertLetter)
        Me.grpRpt_CoveringLetter.Items.Add(Me.grpLetter_insertMemo)
        Me.grpRpt_CoveringLetter.Items.Add(Me.grpCoveringLetter_mnu6)
        Me.grpRpt_CoveringLetter.Items.Add(Me.mnuCloseDocuments11)
        Me.grpRpt_CoveringLetter.Items.Add(Me.grpLetter_delReport)
        Me.grpRpt_CoveringLetter.Label = "Templates"
        Me.grpRpt_CoveringLetter.Name = "grpRpt_CoveringLetter"
        '
        'grpLetter_insertLetter
        '
        Me.grpLetter_insertLetter.KeyTip = "LT"
        Me.grpLetter_insertLetter.Label = "Letter"
        Me.grpLetter_insertLetter.Name = "grpLetter_insertLetter"
        Me.grpLetter_insertLetter.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.grpLetter_insertLetter.ScreenTip = "Letter"
        Me.grpLetter_insertLetter.ShowImage = True
        Me.grpLetter_insertLetter.SuperTip = "Inserts a formatted letter in its own section at the beginning of the document."
        '
        'grpLetter_insertMemo
        '
        Me.grpLetter_insertMemo.KeyTip = "LM"
        Me.grpLetter_insertMemo.Label = "Memo"
        Me.grpLetter_insertMemo.Name = "grpLetter_insertMemo"
        Me.grpLetter_insertMemo.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.grpLetter_insertMemo.ScreenTip = "Memo"
        Me.grpLetter_insertMemo.ShowImage = True
        Me.grpLetter_insertMemo.SuperTip = "Inserts a formatted memo in its own section at the beginning of the document."
        '
        'grpCoveringLetter_mnu6
        '
        Me.grpCoveringLetter_mnu6.Items.Add(Me.mnuCloseDocuments777)
        Me.grpCoveringLetter_mnu6.Items.Add(Me.Separator61)
        Me.grpCoveringLetter_mnu6.Items.Add(Me.grpLetter_btn_forMemo)
        Me.grpCoveringLetter_mnu6.KeyTip = "LD"
        Me.grpCoveringLetter_mnu6.Label = "Contact details"
        Me.grpCoveringLetter_mnu6.Name = "grpCoveringLetter_mnu6"
        Me.grpCoveringLetter_mnu6.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.grpCoveringLetter_mnu6.ScreenTip = "Contact details"
        Me.grpCoveringLetter_mnu6.ShowImage = True
        Me.grpCoveringLetter_mnu6.SuperTip = "The user can choose what contact details to include at the right hand side of the" &
    " new stationery"
        '
        'mnuCloseDocuments777
        '
        Me.mnuCloseDocuments777.Items.Add(Me.grpLetter_Melbourne)
        Me.mnuCloseDocuments777.Items.Add(Me.grpLetter_Sydney)
        Me.mnuCloseDocuments777.Items.Add(Me.grpLetter_Brisbane)
        Me.mnuCloseDocuments777.Items.Add(Me.grpLetter_Canberra)
        Me.mnuCloseDocuments777.Items.Add(Me.grpLetter_Perth)
        Me.mnuCloseDocuments777.Items.Add(Me.grpLetter_Adelaide)
        Me.mnuCloseDocuments777.Label = "For &Letterhead"
        Me.mnuCloseDocuments777.Name = "mnuCloseDocuments777"
        Me.mnuCloseDocuments777.OfficeImageId = "BevelShapeGallery"
        Me.mnuCloseDocuments777.ScreenTip = "For Letterhead"
        Me.mnuCloseDocuments777.ShowImage = True
        Me.mnuCloseDocuments777.SuperTip = "With your cursor somewhere in the letter page, choose your office from this menu." &
    ""
        '
        'grpLetter_Melbourne
        '
        Me.grpLetter_Melbourne.Label = "&Melbourne"
        Me.grpLetter_Melbourne.Name = "grpLetter_Melbourne"
        Me.grpLetter_Melbourne.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.grpLetter_Melbourne.ScreenTip = "Melbourne"
        Me.grpLetter_Melbourne.ShowImage = True
        Me.grpLetter_Melbourne.SuperTip = "Inserts Melbourne contact details"
        '
        'grpLetter_Sydney
        '
        Me.grpLetter_Sydney.Label = "&Sydney"
        Me.grpLetter_Sydney.Name = "grpLetter_Sydney"
        Me.grpLetter_Sydney.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.grpLetter_Sydney.ScreenTip = "Sydney"
        Me.grpLetter_Sydney.ShowImage = True
        Me.grpLetter_Sydney.SuperTip = "Inserts Sydney contact details"
        '
        'grpLetter_Brisbane
        '
        Me.grpLetter_Brisbane.Label = "&Brisbane"
        Me.grpLetter_Brisbane.Name = "grpLetter_Brisbane"
        Me.grpLetter_Brisbane.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.grpLetter_Brisbane.ScreenTip = "Brisbane"
        Me.grpLetter_Brisbane.ShowImage = True
        Me.grpLetter_Brisbane.SuperTip = "Inserts Brisbane contact details"
        '
        'grpLetter_Canberra
        '
        Me.grpLetter_Canberra.Label = "&Canberra"
        Me.grpLetter_Canberra.Name = "grpLetter_Canberra"
        Me.grpLetter_Canberra.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.grpLetter_Canberra.ScreenTip = "Canberra"
        Me.grpLetter_Canberra.ShowImage = True
        Me.grpLetter_Canberra.SuperTip = "Inserts Canberra contact details"
        '
        'grpLetter_Perth
        '
        Me.grpLetter_Perth.Label = "&Perth"
        Me.grpLetter_Perth.Name = "grpLetter_Perth"
        Me.grpLetter_Perth.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.grpLetter_Perth.ScreenTip = "Perth"
        Me.grpLetter_Perth.ShowImage = True
        Me.grpLetter_Perth.SuperTip = "Inserts Perth contact details"
        '
        'grpLetter_Adelaide
        '
        Me.grpLetter_Adelaide.Label = "&Adelaide"
        Me.grpLetter_Adelaide.Name = "grpLetter_Adelaide"
        Me.grpLetter_Adelaide.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.grpLetter_Adelaide.ScreenTip = "Adelaide"
        Me.grpLetter_Adelaide.ShowImage = True
        Me.grpLetter_Adelaide.SuperTip = "Inserts Adelaide contact details"
        '
        'Separator61
        '
        Me.Separator61.Name = "Separator61"
        '
        'grpLetter_btn_forMemo
        '
        Me.grpLetter_btn_forMemo.Label = "For &Memo"
        Me.grpLetter_btn_forMemo.Name = "grpLetter_btn_forMemo"
        Me.grpLetter_btn_forMemo.OfficeImageId = "BevelShapeGallery"
        Me.grpLetter_btn_forMemo.ShowImage = True
        Me.grpLetter_btn_forMemo.SuperTip = "Inserts contact details for a memo"
        '
        'mnuCloseDocuments11
        '
        Me.mnuCloseDocuments11.Items.Add(Me.grpLetter_LtrHead1)
        Me.mnuCloseDocuments11.Items.Add(Me.grpLetter_LtrHead2)
        Me.mnuCloseDocuments11.Items.Add(Me.grpLetter_LtrHead3)
        Me.mnuCloseDocuments11.KeyTip = "LS"
        Me.mnuCloseDocuments11.Label = "Letter heading styles"
        Me.mnuCloseDocuments11.Name = "mnuCloseDocuments11"
        Me.mnuCloseDocuments11.OfficeImageId = "BevelShapeGallery"
        Me.mnuCloseDocuments11.ScreenTip = "Letter heading styles"
        Me.mnuCloseDocuments11.ShowImage = True
        Me.mnuCloseDocuments11.SuperTip = "Choose heading styles for use within the letter."
        '
        'grpLetter_LtrHead1
        '
        Me.grpLetter_LtrHead1.Label = "Letter Heading &1 (Bold)"
        Me.grpLetter_LtrHead1.Name = "grpLetter_LtrHead1"
        Me.grpLetter_LtrHead1.OfficeImageId = "BevelShapeGallery"
        Me.grpLetter_LtrHead1.ScreenTip = "Bold"
        Me.grpLetter_LtrHead1.ShowImage = True
        '
        'grpLetter_LtrHead2
        '
        Me.grpLetter_LtrHead2.Label = "Letter Heading &2 (Bold Italic)"
        Me.grpLetter_LtrHead2.Name = "grpLetter_LtrHead2"
        Me.grpLetter_LtrHead2.OfficeImageId = "BevelShapeGallery"
        Me.grpLetter_LtrHead2.ScreenTip = "Bold Italic"
        Me.grpLetter_LtrHead2.ShowImage = True
        '
        'grpLetter_LtrHead3
        '
        Me.grpLetter_LtrHead3.Label = "Letter Heading &3 (Italic)"
        Me.grpLetter_LtrHead3.Name = "grpLetter_LtrHead3"
        Me.grpLetter_LtrHead3.OfficeImageId = "BevelShapeGallery"
        Me.grpLetter_LtrHead3.ScreenTip = "Italic"
        Me.grpLetter_LtrHead3.ShowImage = True
        '
        'grpLetter_delReport
        '
        Me.grpLetter_delReport.KeyTip = "LX"
        Me.grpLetter_delReport.Label = "Delete Report"
        Me.grpLetter_delReport.Name = "grpLetter_delReport"
        Me.grpLetter_delReport.OfficeImageId = "DeleteWeb"
        Me.grpLetter_delReport.ScreenTip = "Delete Report"
        Me.grpLetter_delReport.ShowImage = True
        Me.grpLetter_delReport.SuperTip = "This function will delete all document sections after the first section, which is" &
    " only useful if you have attached a letter or Memo to an existing report."
        '
        'grp_WhatsNew
        '
        Me.grp_WhatsNew.Items.Add(Me.grpWhatsNew_Form)
        Me.grp_WhatsNew.Label = "What's new"
        Me.grp_WhatsNew.Name = "grp_WhatsNew"
        '
        'grpWhatsNew_Form
        '
        Me.grpWhatsNew_Form.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grpWhatsNew_Form.Label = "What's new ?"
        Me.grpWhatsNew_Form.Name = "grpWhatsNew_Form"
        Me.grpWhatsNew_Form.OfficeImageId = "NewCustomButton"
        Me.grpWhatsNew_Form.ShowImage = True
        Me.grpWhatsNew_Form.SuperTip = resources.GetString("grpWhatsNew_Form.SuperTip")
        '
        'grp_Fixes
        '
        Me.grp_Fixes.Items.Add(Me.grpFixes_Repairs)
        Me.grp_Fixes.Items.Add(Me.mnu_Pagination)
        Me.grp_Fixes.Items.Add(Me.grpFixes_mnu_Other)
        Me.grp_Fixes.Label = "Fixes"
        Me.grp_Fixes.Name = "grp_Fixes"
        '
        'grpFixes_Repairs
        '
        Me.grpFixes_Repairs.Items.Add(Me.grpFixes_Repairs_remCharChar)
        Me.grpFixes_Repairs.Items.Add(Me.grpFixes_Repairs_remSpaces_indrCells)
        Me.grpFixes_Repairs.Items.Add(Me.grpFixes_Repairs_SetLanguage)
        Me.grpFixes_Repairs.KeyTip = "XR"
        Me.grpFixes_Repairs.Label = "Repairs"
        Me.grpFixes_Repairs.Name = "grpFixes_Repairs"
        Me.grpFixes_Repairs.OfficeImageId = "FileStartWorkflow"
        Me.grpFixes_Repairs.ShowImage = True
        '
        'grpFixes_Repairs_remCharChar
        '
        Me.grpFixes_Repairs_remCharChar.Label = "&Remove char char"
        Me.grpFixes_Repairs_remCharChar.Name = "grpFixes_Repairs_remCharChar"
        Me.grpFixes_Repairs_remCharChar.OfficeImageId = "HappyFace"
        Me.grpFixes_Repairs_remCharChar.ScreenTip = "Remove char char"
        Me.grpFixes_Repairs_remCharChar.ShowImage = True
        Me.grpFixes_Repairs_remCharChar.SuperTip = "Removes char char infestation and style aliases - leaving clean style names.  If " &
    "you are having problems running the Custom Table Macro, run this macro first, th" &
    "en re-run the Table macro."
        '
        'grpFixes_Repairs_remSpaces_indrCells
        '
        Me.grpFixes_Repairs_remSpaces_indrCells.Label = "&Trim white spaces in tables"
        Me.grpFixes_Repairs_remSpaces_indrCells.Name = "grpFixes_Repairs_remSpaces_indrCells"
        Me.grpFixes_Repairs_remSpaces_indrCells.OfficeImageId = "DeleteSpaces"
        Me.grpFixes_Repairs_remSpaces_indrCells.ScreenTip = "Trim white spaces in tables"
        Me.grpFixes_Repairs_remSpaces_indrCells.ShowImage = True
        Me.grpFixes_Repairs_remSpaces_indrCells.SuperTip = resources.GetString("grpFixes_Repairs_remSpaces_indrCells.SuperTip")
        '
        'grpFixes_Repairs_SetLanguage
        '
        Me.grpFixes_Repairs_SetLanguage.Label = "Change language to English (Australia)"
        Me.grpFixes_Repairs_SetLanguage.Name = "grpFixes_Repairs_SetLanguage"
        Me.grpFixes_Repairs_SetLanguage.OfficeImageId = "SetLanguage"
        Me.grpFixes_Repairs_SetLanguage.ScreenTip = "English (Australian)"
        Me.grpFixes_Repairs_SetLanguage.ShowImage = True
        Me.grpFixes_Repairs_SetLanguage.SuperTip = "Changes language of whole document to English (Australia)."
        '
        'mnu_Pagination
        '
        Me.mnu_Pagination.Items.Add(Me.grpFixes_RePaginate)
        Me.mnu_Pagination.Items.Add(Me.grpFixes_PaginateOff)
        Me.mnu_Pagination.KeyTip = "XG"
        Me.mnu_Pagination.Label = "Pagination"
        Me.mnu_Pagination.Name = "mnu_Pagination"
        Me.mnu_Pagination.OfficeImageId = "AccessRefreshAllLists"
        Me.mnu_Pagination.ShowImage = True
        '
        'grpFixes_RePaginate
        '
        Me.grpFixes_RePaginate.Label = "&RePaginate"
        Me.grpFixes_RePaginate.Name = "grpFixes_RePaginate"
        Me.grpFixes_RePaginate.OfficeImageId = "RePaginate"
        Me.grpFixes_RePaginate.ScreenTip = "RePaginate"
        Me.grpFixes_RePaginate.ShowImage = True
        Me.grpFixes_RePaginate.SuperTip = "Repaginate  document."
        '
        'grpFixes_PaginateOff
        '
        Me.grpFixes_PaginateOff.Label = "Turn pagination &off"
        Me.grpFixes_PaginateOff.Name = "grpFixes_PaginateOff"
        Me.grpFixes_PaginateOff.OfficeImageId = "PaginationOff"
        Me.grpFixes_PaginateOff.ScreenTip = "Turn pagination off"
        Me.grpFixes_PaginateOff.ShowImage = True
        Me.grpFixes_PaginateOff.SuperTip = "Stops background  pagination."
        '
        'grpFixes_mnu_Other
        '
        Me.grpFixes_mnu_Other.Items.Add(Me.mnu_Fixes_ScreenUpdating)
        Me.grpFixes_mnu_Other.KeyTip = "XO"
        Me.grpFixes_mnu_Other.Label = "Other Fixes"
        Me.grpFixes_mnu_Other.Name = "grpFixes_mnu_Other"
        Me.grpFixes_mnu_Other.OfficeImageId = "BevelShapeGallery"
        Me.grpFixes_mnu_Other.ScreenTip = "Miscellaneous Fixes"
        Me.grpFixes_mnu_Other.ShowImage = True
        Me.grpFixes_mnu_Other.SuperTip = "This menu contains a variety of fixes"
        '
        'mnu_Fixes_ScreenUpdating
        '
        Me.mnu_Fixes_ScreenUpdating.Items.Add(Me.grpFixes_ScreenUpdatingOff)
        Me.mnu_Fixes_ScreenUpdating.Items.Add(Me.grpFixes_ScreenUpdatingOn)
        Me.mnu_Fixes_ScreenUpdating.Label = "Screen Updating"
        Me.mnu_Fixes_ScreenUpdating.Name = "mnu_Fixes_ScreenUpdating"
        Me.mnu_Fixes_ScreenUpdating.OfficeImageId = "BevelShapeGallery"
        Me.mnu_Fixes_ScreenUpdating.ShowImage = True
        '
        'grpFixes_ScreenUpdatingOff
        '
        Me.grpFixes_ScreenUpdatingOff.Label = "Screen Updating - Off"
        Me.grpFixes_ScreenUpdatingOff.Name = "grpFixes_ScreenUpdatingOff"
        Me.grpFixes_ScreenUpdatingOff.OfficeImageId = "BevelShapeGallery"
        Me.grpFixes_ScreenUpdatingOff.ScreenTip = "Screen Updating - Off"
        Me.grpFixes_ScreenUpdatingOff.ShowImage = True
        Me.grpFixes_ScreenUpdatingOff.SuperTip = "Click this button to stop screen updating.  This may be useful when working with " &
    "track changes in a long or complex document."
        '
        'grpFixes_ScreenUpdatingOn
        '
        Me.grpFixes_ScreenUpdatingOn.Label = "Screen Updating - On"
        Me.grpFixes_ScreenUpdatingOn.Name = "grpFixes_ScreenUpdatingOn"
        Me.grpFixes_ScreenUpdatingOn.OfficeImageId = "BevelShapeGallery"
        Me.grpFixes_ScreenUpdatingOn.ScreenTip = "Screen Updating - On"
        Me.grpFixes_ScreenUpdatingOn.ShowImage = True
        Me.grpFixes_ScreenUpdatingOn.SuperTip = "Click this button to start screen updating."
        '
        'tab_aa_Finalise
        '
        Me.tab_aa_Finalise.Groups.Add(Me.grp_WaterMarks)
        Me.tab_aa_Finalise.Groups.Add(Me.grp_PgNumMgmnt)
        Me.tab_aa_Finalise.Groups.Add(Me.grp_Finalise)
        Me.tab_aa_Finalise.Groups.Add(Me.grpWCAG)
        Me.tab_aa_Finalise.Groups.Add(Me.grpRbn_Mgmnt)
        Me.tab_aa_Finalise.Groups.Add(Me.grpTst_LoadFromWeb)
        Me.tab_aa_Finalise.Groups.Add(Me.grpMetaData)
        Me.tab_aa_Finalise.Groups.Add(Me.grpTestTools)
        Me.tab_aa_Finalise.KeyTip = "JF"
        Me.tab_aa_Finalise.Label = "Finalise"
        Me.tab_aa_Finalise.Name = "tab_aa_Finalise"
        Me.tab_aa_Finalise.Position = Me.Factory.RibbonPosition.BeforeOfficeId("TabInsert")
        '
        'grp_WaterMarks
        '
        Me.grp_WaterMarks.Items.Add(Me.grp_waterMark_mnu03)
        Me.grp_WaterMarks.Items.Add(Me.grp_waterMark_mnu01)
        Me.grp_WaterMarks.Items.Add(Me.grp_waterMark_mnu02)
        Me.grp_WaterMarks.KeyTip = "N"
        Me.grp_WaterMarks.Label = "WaterMarks"
        Me.grp_WaterMarks.Name = "grp_WaterMarks"
        '
        'grp_waterMark_mnu03
        '
        Me.grp_waterMark_mnu03.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grp_waterMark_mnu03.Items.Add(Me.grp_waterMark_cabinet_add)
        Me.grp_waterMark_mnu03.Items.Add(Me.grp_waterMark_commercial_add)
        Me.grp_waterMark_mnu03.Items.Add(Me.grp_waterMark_confidential_add)
        Me.grp_waterMark_mnu03.Items.Add(Me.grp_waterMark_restricted_add)
        Me.grp_waterMark_mnu03.Items.Add(Me.Separator84)
        Me.grp_waterMark_mnu03.Items.Add(Me.grp_waterMark_atg_UNOFFICIAL_add)
        Me.grp_waterMark_mnu03.Items.Add(Me.grp_waterMark_atg_OFFICIAL_add)
        Me.grp_waterMark_mnu03.Items.Add(Me.grp_waterMark_atg_OFFICIAL_Sensitive_add)
        Me.grp_waterMark_mnu03.Items.Add(Me.Separator83)
        Me.grp_waterMark_mnu03.Items.Add(Me.grp_waterMark_submnu01)
        Me.grp_waterMark_mnu03.Items.Add(Me.grp_waterMark_colour_red_sec)
        Me.grp_waterMark_mnu03.Items.Add(Me.grp_waterMark_colour_grey_sec)
        Me.grp_waterMark_mnu03.Items.Add(Me.Separator82)
        Me.grp_waterMark_mnu03.Items.Add(Me.grp_waterMark_alignment_Centre_sec)
        Me.grp_waterMark_mnu03.Items.Add(Me.grp_waterMark_alignment_Right_sec)
        Me.grp_waterMark_mnu03.Items.Add(Me.Separator81)
        Me.grp_waterMark_mnu03.Items.Add(Me.grp_waterMark_forceSec_StyleToDefault)
        Me.grp_waterMark_mnu03.KeyTip = "NWS"
        Me.grp_waterMark_mnu03.Label = "Security Level"
        Me.grp_waterMark_mnu03.Name = "grp_waterMark_mnu03"
        Me.grp_waterMark_mnu03.OfficeImageId = "CreateModule"
        Me.grp_waterMark_mnu03.ScreenTip = "Security Level"
        Me.grp_waterMark_mnu03.ShowImage = True
        Me.grp_waterMark_mnu03.SuperTip = "These water marks allow you to mark your document with a specific Security Level." &
    ""
        '
        'grp_waterMark_cabinet_add
        '
        Me.grp_waterMark_cabinet_add.Label = "Add 'Ca&binet-in-Confidence' Watermark"
        Me.grp_waterMark_cabinet_add.Name = "grp_waterMark_cabinet_add"
        Me.grp_waterMark_cabinet_add.OfficeImageId = "AddAccount"
        Me.grp_waterMark_cabinet_add.ScreenTip = "Cabinet-in-Confidence"
        Me.grp_waterMark_cabinet_add.ShowImage = True
        Me.grp_waterMark_cabinet_add.SuperTip = "This button will add the custom 'Cabinet-in-Confidence' watermark"
        '
        'grp_waterMark_commercial_add
        '
        Me.grp_waterMark_commercial_add.Label = "Add 'Co&mmercial-in-Confidence' Watermark"
        Me.grp_waterMark_commercial_add.Name = "grp_waterMark_commercial_add"
        Me.grp_waterMark_commercial_add.OfficeImageId = "AddAccount"
        Me.grp_waterMark_commercial_add.ScreenTip = "Commercial-in-Confidence"
        Me.grp_waterMark_commercial_add.ShowImage = True
        Me.grp_waterMark_commercial_add.SuperTip = "This button will add the custom 'Commercial-in-Confidence' watermark"
        '
        'grp_waterMark_confidential_add
        '
        Me.grp_waterMark_confidential_add.Label = "Add 'Con&fidential' Watermark"
        Me.grp_waterMark_confidential_add.Name = "grp_waterMark_confidential_add"
        Me.grp_waterMark_confidential_add.OfficeImageId = "AddAccount"
        Me.grp_waterMark_confidential_add.ScreenTip = "Confidential"
        Me.grp_waterMark_confidential_add.ShowImage = True
        Me.grp_waterMark_confidential_add.SuperTip = "This button will add the custom 'Confidential' watermark"
        '
        'grp_waterMark_restricted_add
        '
        Me.grp_waterMark_restricted_add.Label = "Add '&Restricted circulation' Watermark"
        Me.grp_waterMark_restricted_add.Name = "grp_waterMark_restricted_add"
        Me.grp_waterMark_restricted_add.OfficeImageId = "AddAccount"
        Me.grp_waterMark_restricted_add.ScreenTip = "Restricted circulation"
        Me.grp_waterMark_restricted_add.ShowImage = True
        Me.grp_waterMark_restricted_add.SuperTip = "This button will add the custom 'Restricted circulation' watermark"
        '
        'Separator84
        '
        Me.Separator84.Name = "Separator84"
        '
        'grp_waterMark_atg_UNOFFICIAL_add
        '
        Me.grp_waterMark_atg_UNOFFICIAL_add.Label = "Add '&UNOFFICIAL' Watermark"
        Me.grp_waterMark_atg_UNOFFICIAL_add.Name = "grp_waterMark_atg_UNOFFICIAL_add"
        Me.grp_waterMark_atg_UNOFFICIAL_add.OfficeImageId = "AddAccount"
        Me.grp_waterMark_atg_UNOFFICIAL_add.ScreenTip = "UNOFFICIAL"
        Me.grp_waterMark_atg_UNOFFICIAL_add.ShowImage = True
        Me.grp_waterMark_atg_UNOFFICIAL_add.SuperTip = "This button will add the custom 'UNOFFICIAL' watermark"
        '
        'grp_waterMark_atg_OFFICIAL_add
        '
        Me.grp_waterMark_atg_OFFICIAL_add.Label = "Add '&OFFICIAL' Watermark"
        Me.grp_waterMark_atg_OFFICIAL_add.Name = "grp_waterMark_atg_OFFICIAL_add"
        Me.grp_waterMark_atg_OFFICIAL_add.OfficeImageId = "AddAccount"
        Me.grp_waterMark_atg_OFFICIAL_add.ScreenTip = "OFFICIAL"
        Me.grp_waterMark_atg_OFFICIAL_add.ShowImage = True
        Me.grp_waterMark_atg_OFFICIAL_add.SuperTip = "This button will add the custom 'OFFICIAL' watermark"
        '
        'grp_waterMark_atg_OFFICIAL_Sensitive_add
        '
        Me.grp_waterMark_atg_OFFICIAL_Sensitive_add.Label = "Add 'OFFICIAL:Sensitive' Watermark"
        Me.grp_waterMark_atg_OFFICIAL_Sensitive_add.Name = "grp_waterMark_atg_OFFICIAL_Sensitive_add"
        Me.grp_waterMark_atg_OFFICIAL_Sensitive_add.OfficeImageId = "AddAccount"
        Me.grp_waterMark_atg_OFFICIAL_Sensitive_add.ScreenTip = "OFFICIAL:Sensitive"
        Me.grp_waterMark_atg_OFFICIAL_Sensitive_add.ShowImage = True
        Me.grp_waterMark_atg_OFFICIAL_Sensitive_add.SuperTip = "This button will add the custom 'OFFICIAL:Sensitive' watermark"
        '
        'Separator83
        '
        Me.Separator83.Name = "Separator83"
        '
        'grp_waterMark_submnu01
        '
        Me.grp_waterMark_submnu01.Items.Add(Me.grp_waterMark_bold_sec)
        Me.grp_waterMark_submnu01.Items.Add(Me.grp_waterMark_NOTbold_sec)
        Me.grp_waterMark_submnu01.Label = "Set to &bold or not bold"
        Me.grp_waterMark_submnu01.Name = "grp_waterMark_submnu01"
        Me.grp_waterMark_submnu01.OfficeImageId = "BorderStyle"
        Me.grp_waterMark_submnu01.ScreenTip = "Bold ot not bold"
        Me.grp_waterMark_submnu01.ShowImage = True
        Me.grp_waterMark_submnu01.SuperTip = "This menu item allows you set the watermark to bold or not bold."
        '
        'grp_waterMark_bold_sec
        '
        Me.grp_waterMark_bold_sec.Label = "Set Water Mark to &bold"
        Me.grp_waterMark_bold_sec.Name = "grp_waterMark_bold_sec"
        Me.grp_waterMark_bold_sec.OfficeImageId = "Bold"
        Me.grp_waterMark_bold_sec.ScreenTip = "Bold"
        Me.grp_waterMark_bold_sec.ShowImage = True
        Me.grp_waterMark_bold_sec.SuperTip = "This button will chnage the watermark text to bold."
        '
        'grp_waterMark_NOTbold_sec
        '
        Me.grp_waterMark_NOTbold_sec.Label = "Set Water Mark to &not bold"
        Me.grp_waterMark_NOTbold_sec.Name = "grp_waterMark_NOTbold_sec"
        Me.grp_waterMark_NOTbold_sec.OfficeImageId = "ClearFormatting"
        Me.grp_waterMark_NOTbold_sec.ScreenTip = "Not bold"
        Me.grp_waterMark_NOTbold_sec.ShowImage = True
        Me.grp_waterMark_NOTbold_sec.SuperTip = "This button will change the watermark text to 'not bold'."
        '
        'grp_waterMark_colour_red_sec
        '
        Me.grp_waterMark_colour_red_sec.Label = "Change colour to &Red"
        Me.grp_waterMark_colour_red_sec.Name = "grp_waterMark_colour_red_sec"
        Me.grp_waterMark_colour_red_sec.OfficeImageId = "ColorRed"
        Me.grp_waterMark_colour_red_sec.ScreenTip = "Red"
        Me.grp_waterMark_colour_red_sec.ShowImage = True
        Me.grp_waterMark_colour_red_sec.SuperTip = "This button will change existing security level watermarks to red."
        '
        'grp_waterMark_colour_grey_sec
        '
        Me.grp_waterMark_colour_grey_sec.Label = "Change colour to &Grey"
        Me.grp_waterMark_colour_grey_sec.Name = "grp_waterMark_colour_grey_sec"
        Me.grp_waterMark_colour_grey_sec.OfficeImageId = "AppointmentColor4"
        Me.grp_waterMark_colour_grey_sec.ScreenTip = "Grey"
        Me.grp_waterMark_colour_grey_sec.ShowImage = True
        Me.grp_waterMark_colour_grey_sec.SuperTip = "This button will change existing security level watermarks to the standard grey (" &
    "default)."
        '
        'Separator82
        '
        Me.Separator82.Name = "Separator82"
        '
        'grp_waterMark_alignment_Centre_sec
        '
        Me.grp_waterMark_alignment_Centre_sec.Label = "Set alignment to the cen&tre"
        Me.grp_waterMark_alignment_Centre_sec.Name = "grp_waterMark_alignment_Centre_sec"
        Me.grp_waterMark_alignment_Centre_sec.OfficeImageId = "AlignCenter"
        Me.grp_waterMark_alignment_Centre_sec.ScreenTip = "Align Centre"
        Me.grp_waterMark_alignment_Centre_sec.ShowImage = True
        Me.grp_waterMark_alignment_Centre_sec.SuperTip = "This button will change the alignment of existing security level watermarks to th" &
    "e centre of the header"
        '
        'grp_waterMark_alignment_Right_sec
        '
        Me.grp_waterMark_alignment_Right_sec.Label = "Set alignment to the r&ight"
        Me.grp_waterMark_alignment_Right_sec.Name = "grp_waterMark_alignment_Right_sec"
        Me.grp_waterMark_alignment_Right_sec.OfficeImageId = "AlignRight"
        Me.grp_waterMark_alignment_Right_sec.ScreenTip = "Align right"
        Me.grp_waterMark_alignment_Right_sec.ShowImage = True
        Me.grp_waterMark_alignment_Right_sec.SuperTip = "This button will change the alignment of existing security level watermarks to th" &
    "e right of the header (default)."
        '
        'Separator81
        '
        Me.Separator81.Name = "Separator81"
        '
        'grp_waterMark_forceSec_StyleToDefault
        '
        Me.grp_waterMark_forceSec_StyleToDefault.Label = "Reset the style to it's &default"
        Me.grp_waterMark_forceSec_StyleToDefault.Name = "grp_waterMark_forceSec_StyleToDefault"
        Me.grp_waterMark_forceSec_StyleToDefault.OfficeImageId = "BevelShapeGallery"
        Me.grp_waterMark_forceSec_StyleToDefault.ScreenTip = "Reset to default"
        Me.grp_waterMark_forceSec_StyleToDefault.ShowImage = True
        Me.grp_waterMark_forceSec_StyleToDefault.SuperTip = "This button will reset the style used for the document security levels to it's de" &
    "fault settings"
        '
        'grp_waterMark_mnu01
        '
        Me.grp_waterMark_mnu01.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grp_waterMark_mnu01.Items.Add(Me.grp_waterMark_removeAll)
        Me.grp_waterMark_mnu01.Items.Add(Me.Separator78)
        Me.grp_waterMark_mnu01.Items.Add(Me.grp_waterMark_mnu04)
        Me.grp_waterMark_mnu01.Items.Add(Me.grp_waterMark_mnu05)
        Me.grp_waterMark_mnu01.KeyTip = "NWX"
        Me.grp_waterMark_mnu01.Label = "Remove"
        Me.grp_waterMark_mnu01.Name = "grp_waterMark_mnu01"
        Me.grp_waterMark_mnu01.OfficeImageId = "WatermarkRemove"
        Me.grp_waterMark_mnu01.ScreenTip = "Remove Watermarks"
        Me.grp_waterMark_mnu01.ShowImage = True
        Me.grp_waterMark_mnu01.SuperTip = "The functions of this menu group will allow you to remove Acil Allen document cla" &
    "ssification watremarks to or from all pages in your document"
        '
        'grp_waterMark_removeAll
        '
        Me.grp_waterMark_removeAll.Label = "Remove &all Watermarks from the current document"
        Me.grp_waterMark_removeAll.Name = "grp_waterMark_removeAll"
        Me.grp_waterMark_removeAll.OfficeImageId = "WatermarkRemove"
        Me.grp_waterMark_removeAll.ScreenTip = "Remove all Watermarks"
        Me.grp_waterMark_removeAll.ShowImage = True
        Me.grp_waterMark_removeAll.SuperTip = "This button will remove ALL Acil Allen custom Water Marks from the current docume" &
    "nt"
        '
        'Separator78
        '
        Me.Separator78.Name = "Separator78"
        '
        'grp_waterMark_mnu04
        '
        Me.grp_waterMark_mnu04.Items.Add(Me.grp_waterMark_removeSec)
        Me.grp_waterMark_mnu04.Items.Add(Me.grp_waterMark_removeStat)
        Me.grp_waterMark_mnu04.Label = "Remove Specific Water Marks form the current &document"
        Me.grp_waterMark_mnu04.Name = "grp_waterMark_mnu04"
        Me.grp_waterMark_mnu04.OfficeImageId = "WatermarkRemove"
        Me.grp_waterMark_mnu04.ScreenTip = "Remove Watermarks from the document"
        Me.grp_waterMark_mnu04.ShowImage = True
        Me.grp_waterMark_mnu04.SuperTip = "This button will remove ALL instances of a selected Water Mark type from the curr" &
    "ent document"
        '
        'grp_waterMark_removeSec
        '
        Me.grp_waterMark_removeSec.Label = "Remove ALL '&Security Level' Watermarks"
        Me.grp_waterMark_removeSec.Name = "grp_waterMark_removeSec"
        Me.grp_waterMark_removeSec.OfficeImageId = "WatermarkRemove"
        Me.grp_waterMark_removeSec.ScreenTip = "'Security Level' Watermarks"
        Me.grp_waterMark_removeSec.ShowImage = True
        Me.grp_waterMark_removeSec.SuperTip = "Remove ALL 'Security Level' Water Marks form the current document"
        '
        'grp_waterMark_removeStat
        '
        Me.grp_waterMark_removeStat.Label = "Remove ALL 'S&tatus Level' Watermarks"
        Me.grp_waterMark_removeStat.Name = "grp_waterMark_removeStat"
        Me.grp_waterMark_removeStat.OfficeImageId = "WatermarkRemove"
        Me.grp_waterMark_removeStat.ScreenTip = "'Status Level' Watermarks"
        Me.grp_waterMark_removeStat.ShowImage = True
        Me.grp_waterMark_removeStat.SuperTip = "This button will remove ALL 'Status Level' Water Mark(s) from the entire document" &
    ""
        '
        'grp_waterMark_mnu05
        '
        Me.grp_waterMark_mnu05.Items.Add(Me.grp_waterMark_removeSec_fromSect)
        Me.grp_waterMark_mnu05.Items.Add(Me.grp_waterMark_removeStat_fromSect)
        Me.grp_waterMark_mnu05.Label = "Remove Specific Water Marks from the current &section"
        Me.grp_waterMark_mnu05.Name = "grp_waterMark_mnu05"
        Me.grp_waterMark_mnu05.OfficeImageId = "WatermarkRemove"
        Me.grp_waterMark_mnu05.ScreenTip = "Remove Watermarks from the section"
        Me.grp_waterMark_mnu05.ShowImage = True
        Me.grp_waterMark_mnu05.SuperTip = "Remove Specific Water Marks from the current section"
        '
        'grp_waterMark_removeSec_fromSect
        '
        Me.grp_waterMark_removeSec_fromSect.Label = "Remove section '&Security Level' Watermarks"""
        Me.grp_waterMark_removeSec_fromSect.Name = "grp_waterMark_removeSec_fromSect"
        Me.grp_waterMark_removeSec_fromSect.OfficeImageId = "WatermarkRemove"
        Me.grp_waterMark_removeSec_fromSect.ScreenTip = "Remove Security Level"
        Me.grp_waterMark_removeSec_fromSect.ShowImage = True
        Me.grp_waterMark_removeSec_fromSect.SuperTip = "his button will remove any 'Security Level' Water Mark(s) from the current sectio" &
    "n of the document (i.e. the section containing your cursor)."
        '
        'grp_waterMark_removeStat_fromSect
        '
        Me.grp_waterMark_removeStat_fromSect.Label = "Remove section 'S&tatus Level' Watermarks"
        Me.grp_waterMark_removeStat_fromSect.Name = "grp_waterMark_removeStat_fromSect"
        Me.grp_waterMark_removeStat_fromSect.OfficeImageId = "WatermarkRemove"
        Me.grp_waterMark_removeStat_fromSect.ScreenTip = "Remove Status Level"
        Me.grp_waterMark_removeStat_fromSect.ShowImage = True
        Me.grp_waterMark_removeStat_fromSect.SuperTip = "This button will remove any 'Status Level' Water Mark(s) from the current section" &
    " of the document (i.e. the section containing your cursor)."
        '
        'grp_waterMark_mnu02
        '
        Me.grp_waterMark_mnu02.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grp_waterMark_mnu02.Items.Add(Me.grp_waterMark_draft_add)
        Me.grp_waterMark_mnu02.Items.Add(Me.grp_waterMark_draftOnly_add)
        Me.grp_waterMark_mnu02.Items.Add(Me.Separator79)
        Me.grp_waterMark_mnu02.Items.Add(Me.grp_waterMark_colour_red_stat)
        Me.grp_waterMark_mnu02.Items.Add(Me.grp_waterMark_colour_grey_stat)
        Me.grp_waterMark_mnu02.Items.Add(Me.Separator80)
        Me.grp_waterMark_mnu02.Items.Add(Me.grp_waterMark_forceStat_StyleToDefault)
        Me.grp_waterMark_mnu02.KeyTip = "NWT"
        Me.grp_waterMark_mnu02.Label = "Release Status"
        Me.grp_waterMark_mnu02.Name = "grp_waterMark_mnu02"
        Me.grp_waterMark_mnu02.OfficeImageId = "CreateModule"
        Me.grp_waterMark_mnu02.ScreenTip = "Release Status"
        Me.grp_waterMark_mnu02.ShowImage = True
        Me.grp_waterMark_mnu02.SuperTip = "These water marks allow you to mark your document with a specific Release Status." &
    ""
        '
        'grp_waterMark_draft_add
        '
        Me.grp_waterMark_draft_add.Label = "Add 'DRA&FT' Watermark"
        Me.grp_waterMark_draft_add.Name = "grp_waterMark_draft_add"
        Me.grp_waterMark_draft_add.OfficeImageId = "AddAccount"
        Me.grp_waterMark_draft_add.ScreenTip = "DRAFT"
        Me.grp_waterMark_draft_add.ShowImage = True
        Me.grp_waterMark_draft_add.SuperTip = "This button will add the custom 'DRAFT' watermark"
        '
        'grp_waterMark_draftOnly_add
        '
        Me.grp_waterMark_draftOnly_add.Label = "Add 'DRAFT &ONLY' Watermark"
        Me.grp_waterMark_draftOnly_add.Name = "grp_waterMark_draftOnly_add"
        Me.grp_waterMark_draftOnly_add.OfficeImageId = "AddAccount"
        Me.grp_waterMark_draftOnly_add.ScreenTip = "DRAFT ONLY"
        Me.grp_waterMark_draftOnly_add.ShowImage = True
        Me.grp_waterMark_draftOnly_add.SuperTip = "This button will add the custom 'DRAFT ONLY' watermark"
        '
        'Separator79
        '
        Me.Separator79.Name = "Separator79"
        '
        'grp_waterMark_colour_red_stat
        '
        Me.grp_waterMark_colour_red_stat.Label = "Change colour to &Red"
        Me.grp_waterMark_colour_red_stat.Name = "grp_waterMark_colour_red_stat"
        Me.grp_waterMark_colour_red_stat.OfficeImageId = "ColorRed"
        Me.grp_waterMark_colour_red_stat.ScreenTip = "Red"
        Me.grp_waterMark_colour_red_stat.ShowImage = True
        Me.grp_waterMark_colour_red_stat.SuperTip = "This button will change existing document status level watermarks to red."
        '
        'grp_waterMark_colour_grey_stat
        '
        Me.grp_waterMark_colour_grey_stat.Label = "Change colour to &Grey"
        Me.grp_waterMark_colour_grey_stat.Name = "grp_waterMark_colour_grey_stat"
        Me.grp_waterMark_colour_grey_stat.OfficeImageId = "AppointmentColor4"
        Me.grp_waterMark_colour_grey_stat.ScreenTip = "Grey"
        Me.grp_waterMark_colour_grey_stat.ShowImage = True
        Me.grp_waterMark_colour_grey_stat.SuperTip = "This button will change existing document status level watermarks to the standard" &
    " grey (default)."
        '
        'Separator80
        '
        Me.Separator80.Name = "Separator80"
        '
        'grp_waterMark_forceStat_StyleToDefault
        '
        Me.grp_waterMark_forceStat_StyleToDefault.Label = "Reset the style to it's &default"
        Me.grp_waterMark_forceStat_StyleToDefault.Name = "grp_waterMark_forceStat_StyleToDefault"
        Me.grp_waterMark_forceStat_StyleToDefault.ScreenTip = "Watermark default style"
        Me.grp_waterMark_forceStat_StyleToDefault.ShowImage = True
        Me.grp_waterMark_forceStat_StyleToDefault.SuperTip = "This button will reset the style used for the document status to it's default set" &
    "tings"
        '
        'grp_PgNumMgmnt
        '
        Me.grp_PgNumMgmnt.Items.Add(Me.tabFin_mnu_PageNumFormatting)
        Me.grp_PgNumMgmnt.Items.Add(Me.tabFin_mnu_PgNumMgmnt_ReNum)
        Me.grp_PgNumMgmnt.Label = "Page Numbering"
        Me.grp_PgNumMgmnt.Name = "grp_PgNumMgmnt"
        '
        'tabFin_mnu_PageNumFormatting
        '
        Me.tabFin_mnu_PageNumFormatting.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.tabFin_mnu_PageNumFormatting.Items.Add(Me.grpFixes_ApplyEsNumbering)
        Me.tabFin_mnu_PageNumFormatting.Items.Add(Me.grpFixes_ApplyStdNumbering)
        Me.tabFin_mnu_PageNumFormatting.Items.Add(Me.grpFixes_ApplyAppNumbering)
        Me.tabFin_mnu_PageNumFormatting.Items.Add(Me.Separator86)
        Me.tabFin_mnu_PageNumFormatting.Items.Add(Me.grpFixes_ContinueNumbering)
        Me.tabFin_mnu_PageNumFormatting.Items.Add(Me.grpFixes_RestartNumbering)
        Me.tabFin_mnu_PageNumFormatting.Items.Add(Me.Separator85)
        Me.tabFin_mnu_PageNumFormatting.Items.Add(Me.grpFixes_getNumberingDialog)
        Me.tabFin_mnu_PageNumFormatting.KeyTip = "PF"
        Me.tabFin_mnu_PageNumFormatting.Label = "Page # Formatting"
        Me.tabFin_mnu_PageNumFormatting.Name = "tabFin_mnu_PageNumFormatting"
        Me.tabFin_mnu_PageNumFormatting.OfficeImageId = "LegalBlackline"
        Me.tabFin_mnu_PageNumFormatting.ScreenTip = "Page # Formatting"
        Me.tabFin_mnu_PageNumFormatting.ShowImage = True
        Me.tabFin_mnu_PageNumFormatting.SuperTip = "This function lets you directly control Page Numbering in the section that curren" &
    "tly contains the cursor"
        '
        'grpFixes_ApplyEsNumbering
        '
        Me.grpFixes_ApplyEsNumbering.Label = "Apply &Executive Summary numbering"
        Me.grpFixes_ApplyEsNumbering.Name = "grpFixes_ApplyEsNumbering"
        Me.grpFixes_ApplyEsNumbering.OfficeImageId = "BevelShapeGallery"
        Me.grpFixes_ApplyEsNumbering.ScreenTip = "Executive Summary Numbering"
        Me.grpFixes_ApplyEsNumbering.ShowImage = True
        Me.grpFixes_ApplyEsNumbering.SuperTip = "Applies a roman numeral page numbering scheme."
        '
        'grpFixes_ApplyStdNumbering
        '
        Me.grpFixes_ApplyStdNumbering.Label = "Apply &Standard page number format"
        Me.grpFixes_ApplyStdNumbering.Name = "grpFixes_ApplyStdNumbering"
        Me.grpFixes_ApplyStdNumbering.OfficeImageId = "BevelShapeGallery"
        Me.grpFixes_ApplyStdNumbering.ScreenTip = "Standard Numbering"
        Me.grpFixes_ApplyStdNumbering.ShowImage = True
        Me.grpFixes_ApplyStdNumbering.SuperTip = "Applies standard page numbering to this section."
        '
        'grpFixes_ApplyAppNumbering
        '
        Me.grpFixes_ApplyAppNumbering.Label = "Apply &Appendix style numbering"
        Me.grpFixes_ApplyAppNumbering.Name = "grpFixes_ApplyAppNumbering"
        Me.grpFixes_ApplyAppNumbering.OfficeImageId = "BevelShapeGallery"
        Me.grpFixes_ApplyAppNumbering.ScreenTip = "Appendix Numbering"
        Me.grpFixes_ApplyAppNumbering.ShowImage = True
        Me.grpFixes_ApplyAppNumbering.SuperTip = resources.GetString("grpFixes_ApplyAppNumbering.SuperTip")
        '
        'Separator86
        '
        Me.Separator86.Name = "Separator86"
        '
        'grpFixes_ContinueNumbering
        '
        Me.grpFixes_ContinueNumbering.Label = "&Continue page numbering"
        Me.grpFixes_ContinueNumbering.Name = "grpFixes_ContinueNumbering"
        Me.grpFixes_ContinueNumbering.OfficeImageId = "BevelShapeGallery"
        Me.grpFixes_ContinueNumbering.ScreenTip = "Continue Numbering"
        Me.grpFixes_ContinueNumbering.ShowImage = True
        Me.grpFixes_ContinueNumbering.SuperTip = "Continues page numbering in this section."
        '
        'grpFixes_RestartNumbering
        '
        Me.grpFixes_RestartNumbering.Label = "&Restart page numbering"
        Me.grpFixes_RestartNumbering.Name = "grpFixes_RestartNumbering"
        Me.grpFixes_RestartNumbering.OfficeImageId = "BevelShapeGallery"
        Me.grpFixes_RestartNumbering.ScreenTip = "Restart (i.e. at 1)"
        Me.grpFixes_RestartNumbering.ShowImage = True
        Me.grpFixes_RestartNumbering.SuperTip = "Restarts page numbering in this section."
        '
        'Separator85
        '
        Me.Separator85.Name = "Separator85"
        '
        'grpFixes_getNumberingDialog
        '
        Me.grpFixes_getNumberingDialog.Label = "Display the page numbering &dialog"
        Me.grpFixes_getNumberingDialog.Name = "grpFixes_getNumberingDialog"
        Me.grpFixes_getNumberingDialog.OfficeImageId = "BevelShapeGallery"
        Me.grpFixes_getNumberingDialog.ScreenTip = "Show Standard Dialog"
        Me.grpFixes_getNumberingDialog.ShowImage = True
        Me.grpFixes_getNumberingDialog.SuperTip = "Use this dialog for full control of the numbering scheme of your report. Remember" &
    " that page number adjustments are section specific"
        '
        'tabFin_mnu_PgNumMgmnt_ReNum
        '
        Me.tabFin_mnu_PgNumMgmnt_ReNum.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.tabFin_mnu_PgNumMgmnt_ReNum.Items.Add(Me.grp_PgNumMgmnt_ReNum_std)
        Me.tabFin_mnu_PgNumMgmnt_ReNum.Items.Add(Me.grp_PgNumMgmnt_ReNum_2Part)
        Me.tabFin_mnu_PgNumMgmnt_ReNum.KeyTip = "PR"
        Me.tabFin_mnu_PgNumMgmnt_ReNum.Label = "Renumber Report/Brief"
        Me.tabFin_mnu_PgNumMgmnt_ReNum.Name = "tabFin_mnu_PgNumMgmnt_ReNum"
        Me.tabFin_mnu_PgNumMgmnt_ReNum.OfficeImageId = "PageMenu"
        Me.tabFin_mnu_PgNumMgmnt_ReNum.ScreenTip = "Renumber Report/Brief"
        Me.tabFin_mnu_PgNumMgmnt_ReNum.ShowImage = True
        Me.tabFin_mnu_PgNumMgmnt_ReNum.SuperTip = "This function will renumber the body of the current Report or Brief."
        '
        'grp_PgNumMgmnt_ReNum_std
        '
        Me.grp_PgNumMgmnt_ReNum_std.Label = "Standard page numbers (pageNum)"
        Me.grp_PgNumMgmnt_ReNum_std.Name = "grp_PgNumMgmnt_ReNum_std"
        Me.grp_PgNumMgmnt_ReNum_std.OfficeImageId = "PageMenu"
        Me.grp_PgNumMgmnt_ReNum_std.ScreenTip = "Standard (one part) page numbers"
        Me.grp_PgNumMgmnt_ReNum_std.ShowImage = True
        Me.grp_PgNumMgmnt_ReNum_std.SuperTip = "Check here to ensure the body of this report is numbered using a sequential page " &
    "number format (page number only)"
        '
        'grp_PgNumMgmnt_ReNum_2Part
        '
        Me.grp_PgNumMgmnt_ReNum_2Part.Label = "Two part page numbers (Chpt-pageNum)"
        Me.grp_PgNumMgmnt_ReNum_2Part.Name = "grp_PgNumMgmnt_ReNum_2Part"
        Me.grp_PgNumMgmnt_ReNum_2Part.OfficeImageId = "PageMenu"
        Me.grp_PgNumMgmnt_ReNum_2Part.ScreenTip = "Two part numbering (Chpt-pageNum)"
        Me.grp_PgNumMgmnt_ReNum_2Part.ShowImage = True
        Me.grp_PgNumMgmnt_ReNum_2Part.SuperTip = "Check here to ensure the body of this report is numbered using a 'Chapter-PageNum" &
    "' page number format"
        '
        'grp_Finalise
        '
        Me.grp_Finalise.Items.Add(Me.grp_Finalise_mnu01)
        Me.grp_Finalise.Label = "Finalise"
        Me.grp_Finalise.Name = "grp_Finalise"
        '
        'grp_Finalise_mnu01
        '
        Me.grp_Finalise_mnu01.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grp_Finalise_mnu01.Items.Add(Me.grp_Finalise_CrossRefError)
        Me.grp_Finalise_mnu01.Items.Add(Me.Separator88)
        Me.grp_Finalise_mnu01.Items.Add(Me.grp_Finalise_DoAll)
        Me.grp_Finalise_mnu01.Items.Add(Me.grp_Finalise_AllFunctions)
        Me.grp_Finalise_mnu01.KeyTip = "FF"
        Me.grp_Finalise_mnu01.Label = "Finalise"
        Me.grp_Finalise_mnu01.Name = "grp_Finalise_mnu01"
        Me.grp_Finalise_mnu01.OfficeImageId = "AcceptTask"
        Me.grp_Finalise_mnu01.ScreenTip = "Finalise"
        Me.grp_Finalise_mnu01.ShowImage = True
        Me.grp_Finalise_mnu01.SuperTip = "This menu provides a number of tools that help you finalise a document"
        '
        'grp_Finalise_CrossRefError
        '
        Me.grp_Finalise_CrossRefError.Label = "Check for &cross reference errors"
        Me.grp_Finalise_CrossRefError.Name = "grp_Finalise_CrossRefError"
        Me.grp_Finalise_CrossRefError.OfficeImageId = "AcceptTask"
        Me.grp_Finalise_CrossRefError.ScreenTip = "Find cross reference errors"
        Me.grp_Finalise_CrossRefError.ShowImage = True
        Me.grp_Finalise_CrossRefError.SuperTip = resources.GetString("grp_Finalise_CrossRefError.SuperTip")
        '
        'Separator88
        '
        Me.Separator88.Name = "Separator88"
        '
        'grp_Finalise_DoAll
        '
        Me.grp_Finalise_DoAll.Label = "Do &all of the document 'finishing' functions"
        Me.grp_Finalise_DoAll.Name = "grp_Finalise_DoAll"
        Me.grp_Finalise_DoAll.OfficeImageId = "AcceptTask"
        Me.grp_Finalise_DoAll.ScreenTip = "Do all"
        Me.grp_Finalise_DoAll.ShowImage = True
        Me.grp_Finalise_DoAll.SuperTip = "This function will, sequentially do all of the document 'finishing' functions"
        '
        'grp_Finalise_AllFunctions
        '
        Me.grp_Finalise_AllFunctions.Items.Add(Me.grp_Finalise_upDateCopyrightNotice)
        Me.grp_Finalise_AllFunctions.Items.Add(Me.Separator87)
        Me.grp_Finalise_AllFunctions.Items.Add(Me.grp_Finalise_updateFields)
        Me.grp_Finalise_AllFunctions.Items.Add(Me.grp_Finalise_setFootersToBold)
        Me.grp_Finalise_AllFunctions.Items.Add(Me.grp_Finalise_RefreshTOC)
        Me.grp_Finalise_AllFunctions.Items.Add(Me.grp_Finalise_CrossRefFlds_setToArialNarrow)
        Me.grp_Finalise_AllFunctions.Items.Add(Me.grp_Finalise_CrossRefFlds_setRefFldNotBold)
        Me.grp_Finalise_AllFunctions.Items.Add(Me.grpFixes_Repairs_delSpace1_betweenWords)
        Me.grp_Finalise_AllFunctions.Items.Add(Me.grpFixes_Repairs_delSpace1_atSentenceEnd)
        Me.grp_Finalise_AllFunctions.Label = "&Individual document 'finishing' functions"
        Me.grp_Finalise_AllFunctions.Name = "grp_Finalise_AllFunctions"
        Me.grp_Finalise_AllFunctions.OfficeImageId = "AcceptTask"
        Me.grp_Finalise_AllFunctions.ScreenTip = "Individual finishing functions"
        Me.grp_Finalise_AllFunctions.ShowImage = True
        Me.grp_Finalise_AllFunctions.SuperTip = "This menu provides access to all of the individual document 'finishing' functions" &
    ""
        '
        'grp_Finalise_upDateCopyrightNotice
        '
        Me.grp_Finalise_upDateCopyrightNotice.Label = "Update &Copyright Notice"
        Me.grp_Finalise_upDateCopyrightNotice.Name = "grp_Finalise_upDateCopyrightNotice"
        Me.grp_Finalise_upDateCopyrightNotice.OfficeImageId = "BevelShapeGallery"
        Me.grp_Finalise_upDateCopyrightNotice.ScreenTip = "Copyright notice"
        Me.grp_Finalise_upDateCopyrightNotice.ShowImage = True
        Me.grp_Finalise_upDateCopyrightNotice.SuperTip = "This function will ensure that the copyright notice is current."
        '
        'Separator87
        '
        Me.Separator87.Name = "Separator87"
        '
        'grp_Finalise_updateFields
        '
        Me.grp_Finalise_updateFields.Label = "&Update document fields"
        Me.grp_Finalise_updateFields.Name = "grp_Finalise_updateFields"
        Me.grp_Finalise_updateFields.OfficeImageId = "BevelShapeGallery"
        Me.grp_Finalise_updateFields.ScreenTip = "Update document fields"
        Me.grp_Finalise_updateFields.ShowImage = True
        Me.grp_Finalise_updateFields.SuperTip = "This function will update all fields in the document"
        '
        'grp_Finalise_setFootersToBold
        '
        Me.grp_Finalise_setFootersToBold.Label = "Set Report Footers to &Bold"
        Me.grp_Finalise_setFootersToBold.Name = "grp_Finalise_setFootersToBold"
        Me.grp_Finalise_setFootersToBold.OfficeImageId = "Bold"
        Me.grp_Finalise_setFootersToBold.ScreenTip = "Footers to Bold"
        Me.grp_Finalise_setFootersToBold.ShowImage = True
        Me.grp_Finalise_setFootersToBold.SuperTip = "This function will ensure that all of the footers in your report are set to bold"
        '
        'grp_Finalise_RefreshTOC
        '
        Me.grp_Finalise_RefreshTOC.Label = "Refresh &TOC"
        Me.grp_Finalise_RefreshTOC.Name = "grp_Finalise_RefreshTOC"
        Me.grp_Finalise_RefreshTOC.OfficeImageId = "TableOfContentsGallery"
        Me.grp_Finalise_RefreshTOC.ScreenTip = "Refresh TOC"
        Me.grp_Finalise_RefreshTOC.ShowImage = True
        Me.grp_Finalise_RefreshTOC.SuperTip = "Will refresh the Table of Contents (including the Table of Figures, Tables and Bo" &
    "xes).Depending on the machine and document this can take (on average) anywhere f" &
    "rom 1 second to as much as a minute"
        '
        'grp_Finalise_CrossRefFlds_setToArialNarrow
        '
        Me.grp_Finalise_CrossRefFlds_setToArialNarrow.Label = "Set All Cross Reference Fields to '&Arial Narrow'"
        Me.grp_Finalise_CrossRefFlds_setToArialNarrow.Name = "grp_Finalise_CrossRefFlds_setToArialNarrow"
        Me.grp_Finalise_CrossRefFlds_setToArialNarrow.OfficeImageId = "Repaginate"
        Me.grp_Finalise_CrossRefFlds_setToArialNarrow.ScreenTip = "Set All Cross Reference Fields to 'Arial Narrow'"
        Me.grp_Finalise_CrossRefFlds_setToArialNarrow.ShowImage = True
        Me.grp_Finalise_CrossRefFlds_setToArialNarrow.SuperTip = "This function will ensure that the font of all cross reference fields in the docu" &
    "ment are consistent with the 'Body Text', that is 'Arial Narrow'"
        '
        'grp_Finalise_CrossRefFlds_setRefFldNotBold
        '
        Me.grp_Finalise_CrossRefFlds_setRefFldNotBold.Label = "Set All Cross Reference Fields to '&Not Bold'"
        Me.grp_Finalise_CrossRefFlds_setRefFldNotBold.Name = "grp_Finalise_CrossRefFlds_setRefFldNotBold"
        Me.grp_Finalise_CrossRefFlds_setRefFldNotBold.OfficeImageId = "Repaginate"
        Me.grp_Finalise_CrossRefFlds_setRefFldNotBold.ScreenTip = "Set All Cross Reference Fields to 'Not Bold'"
        Me.grp_Finalise_CrossRefFlds_setRefFldNotBold.ShowImage = True
        Me.grp_Finalise_CrossRefFlds_setRefFldNotBold.SuperTip = "This function will ensure that all cross reference fields in the document are 'No" &
    "t Bold'"
        '
        'grpFixes_Repairs_delSpace1_betweenWords
        '
        Me.grpFixes_Repairs_delSpace1_betweenWords.Label = "1 space only between &words"
        Me.grpFixes_Repairs_delSpace1_betweenWords.Name = "grpFixes_Repairs_delSpace1_betweenWords"
        Me.grpFixes_Repairs_delSpace1_betweenWords.OfficeImageId = "SizeSpaceMenu"
        Me.grpFixes_Repairs_delSpace1_betweenWords.ScreenTip = "Applies to whole document: Searches for 2 spaces between words and replaces them " &
    "with 1. This method would be useful after accepting all tracked changes as part " &
    "of any clean up."
        Me.grpFixes_Repairs_delSpace1_betweenWords.ShowImage = True
        '
        'grpFixes_Repairs_delSpace1_atSentenceEnd
        '
        Me.grpFixes_Repairs_delSpace1_atSentenceEnd.Label = "1 space after end of all &sentences"
        Me.grpFixes_Repairs_delSpace1_atSentenceEnd.Name = "grpFixes_Repairs_delSpace1_atSentenceEnd"
        Me.grpFixes_Repairs_delSpace1_atSentenceEnd.OfficeImageId = "ScheduledProjectStartDate"
        Me.grpFixes_Repairs_delSpace1_atSentenceEnd.ScreenTip = "1 space after end of all &sentences"
        Me.grpFixes_Repairs_delSpace1_atSentenceEnd.ShowImage = True
        Me.grpFixes_Repairs_delSpace1_atSentenceEnd.SuperTip = "Applies to whole document: At end of sentence, after fullstop, question mark or e" &
    "xclamation mark, replaces with 1 space only."
        '
        'grpWCAG
        '
        Me.grpWCAG.Items.Add(Me.tabFin_mnu_AccessibilityTools)
        Me.grpWCAG.Items.Add(Me.grpReport_PlH_convertToInline_findAllFloatingTables)
        Me.grpWCAG.Label = "Accessible document support"
        Me.grpWCAG.Name = "grpWCAG"
        '
        'tabFin_mnu_AccessibilityTools
        '
        Me.tabFin_mnu_AccessibilityTools.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.tabFin_mnu_AccessibilityTools.Items.Add(Me.grpWCAG_notesOnAccessibility)
        Me.tabFin_mnu_AccessibilityTools.Items.Add(Me.grpWCAG_convertThisDoc)
        Me.tabFin_mnu_AccessibilityTools.Items.Add(Me.Separator89)
        Me.tabFin_mnu_AccessibilityTools.Items.Add(Me.grpWCAG_mnu_ContrastControl)
        Me.tabFin_mnu_AccessibilityTools.Items.Add(Me.grpWCAG_tool_tableHeaderColour_all)
        Me.tabFin_mnu_AccessibilityTools.KeyTip = "AC"
        Me.tabFin_mnu_AccessibilityTools.Label = "Accessibility Tools"
        Me.tabFin_mnu_AccessibilityTools.Name = "tabFin_mnu_AccessibilityTools"
        Me.tabFin_mnu_AccessibilityTools.OfficeImageId = "_3DSurfaceMaterialClassic"
        Me.tabFin_mnu_AccessibilityTools.ScreenTip = "Accessibility Tools"
        Me.tabFin_mnu_AccessibilityTools.ShowImage = True
        Me.tabFin_mnu_AccessibilityTools.SuperTip = resources.GetString("tabFin_mnu_AccessibilityTools.SuperTip")
        '
        'grpWCAG_notesOnAccessibility
        '
        Me.grpWCAG_notesOnAccessibility.Label = "Notes on Accessibility"
        Me.grpWCAG_notesOnAccessibility.Name = "grpWCAG_notesOnAccessibility"
        Me.grpWCAG_notesOnAccessibility.OfficeImageId = "BevelShapeGallery"
        Me.grpWCAG_notesOnAccessibility.ScreenTip = "Notes on Accessibility"
        Me.grpWCAG_notesOnAccessibility.ShowImage = True
        '
        'grpWCAG_convertThisDoc
        '
        Me.grpWCAG_convertThisDoc.Label = "Perform a Basic &Accessibility Conversion"
        Me.grpWCAG_convertThisDoc.Name = "grpWCAG_convertThisDoc"
        Me.grpWCAG_convertThisDoc.OfficeImageId = "BevelShapeGallery"
        Me.grpWCAG_convertThisDoc.ScreenTip = "Perform a basic Accessibility conversion"
        Me.grpWCAG_convertThisDoc.ShowImage = True
        Me.grpWCAG_convertThisDoc.SuperTip = resources.GetString("grpWCAG_convertThisDoc.SuperTip")
        '
        'Separator89
        '
        Me.Separator89.Name = "Separator89"
        '
        'grpWCAG_mnu_ContrastControl
        '
        Me.grpWCAG_mnu_ContrastControl.Items.Add(Me.grpWCAG_mnu_SetTransparency)
        Me.grpWCAG_mnu_ContrastControl.Items.Add(Me.grpWCAG_tool_convertAllStyles_toBlack)
        Me.grpWCAG_mnu_ContrastControl.Label = "&Contrast Control..."
        Me.grpWCAG_mnu_ContrastControl.Name = "grpWCAG_mnu_ContrastControl"
        Me.grpWCAG_mnu_ContrastControl.OfficeImageId = "PictureInsertFromFilePowerPoint"
        Me.grpWCAG_mnu_ContrastControl.ScreenTip = "Contrast Control"
        Me.grpWCAG_mnu_ContrastControl.ShowImage = True
        Me.grpWCAG_mnu_ContrastControl.SuperTip = "This menu item provides a set of tools that allow you to adjust the contrast on a" &
    " page... You can set all text to be black and back panel image"
        '
        'grpWCAG_mnu_SetTransparency
        '
        Me.grpWCAG_mnu_SetTransparency.Items.Add(Me.grpWCAG_mnu_SetTransparency_to_0)
        Me.grpWCAG_mnu_SetTransparency.Items.Add(Me.grpWCAG_mnu_SetTransparency_to_25)
        Me.grpWCAG_mnu_SetTransparency.Items.Add(Me.grpWCAG_mnu_SetTransparency_to_50)
        Me.grpWCAG_mnu_SetTransparency.Items.Add(Me.grpWCAG_mnu_SetTransparency_to_75)
        Me.grpWCAG_mnu_SetTransparency.Items.Add(Me.grpWCAG_mnu_SetTransparency_to_100)
        Me.grpWCAG_mnu_SetTransparency.Label = "Set all image back panel(s) &Transparency..."
        Me.grpWCAG_mnu_SetTransparency.Name = "grpWCAG_mnu_SetTransparency"
        Me.grpWCAG_mnu_SetTransparency.OfficeImageId = "PictureInsertFromFilePowerPoint"
        Me.grpWCAG_mnu_SetTransparency.ScreenTip = "Image Back Panels Transparency"
        Me.grpWCAG_mnu_SetTransparency.ShowImage = True
        Me.grpWCAG_mnu_SetTransparency.SuperTip = resources.GetString("grpWCAG_mnu_SetTransparency.SuperTip")
        '
        'grpWCAG_mnu_SetTransparency_to_0
        '
        Me.grpWCAG_mnu_SetTransparency_to_0.Label = "0% transparent (fully &opaque)"
        Me.grpWCAG_mnu_SetTransparency_to_0.Name = "grpWCAG_mnu_SetTransparency_to_0"
        Me.grpWCAG_mnu_SetTransparency_to_0.OfficeImageId = "BevelShapeGallery"
        Me.grpWCAG_mnu_SetTransparency_to_0.ScreenTip = "Full colour"
        Me.grpWCAG_mnu_SetTransparency_to_0.ShowImage = True
        Me.grpWCAG_mnu_SetTransparency_to_0.SuperTip = "Will set the image back panel to 0% transparent"
        '
        'grpWCAG_mnu_SetTransparency_to_25
        '
        Me.grpWCAG_mnu_SetTransparency_to_25.Label = "&25% transparent"
        Me.grpWCAG_mnu_SetTransparency_to_25.Name = "grpWCAG_mnu_SetTransparency_to_25"
        Me.grpWCAG_mnu_SetTransparency_to_25.OfficeImageId = "BevelShapeGallery"
        Me.grpWCAG_mnu_SetTransparency_to_25.ScreenTip = "25% transparent"
        Me.grpWCAG_mnu_SetTransparency_to_25.ShowImage = True
        Me.grpWCAG_mnu_SetTransparency_to_25.SuperTip = "Will set the image back panel to 25% transparent"
        '
        'grpWCAG_mnu_SetTransparency_to_50
        '
        Me.grpWCAG_mnu_SetTransparency_to_50.Label = "&50% transparent"
        Me.grpWCAG_mnu_SetTransparency_to_50.Name = "grpWCAG_mnu_SetTransparency_to_50"
        Me.grpWCAG_mnu_SetTransparency_to_50.OfficeImageId = "BevelShapeGallery"
        Me.grpWCAG_mnu_SetTransparency_to_50.ScreenTip = "50% transparent"
        Me.grpWCAG_mnu_SetTransparency_to_50.ShowImage = True
        Me.grpWCAG_mnu_SetTransparency_to_50.SuperTip = "Will set the image back panel to 50% transparent"
        '
        'grpWCAG_mnu_SetTransparency_to_75
        '
        Me.grpWCAG_mnu_SetTransparency_to_75.Label = "&75% transparent"
        Me.grpWCAG_mnu_SetTransparency_to_75.Name = "grpWCAG_mnu_SetTransparency_to_75"
        Me.grpWCAG_mnu_SetTransparency_to_75.OfficeImageId = "BevelShapeGallery"
        Me.grpWCAG_mnu_SetTransparency_to_75.ScreenTip = "75% transparent"
        Me.grpWCAG_mnu_SetTransparency_to_75.ShowImage = True
        Me.grpWCAG_mnu_SetTransparency_to_75.SuperTip = "Will set the image back panel to 75% transparent"
        '
        'grpWCAG_mnu_SetTransparency_to_100
        '
        Me.grpWCAG_mnu_SetTransparency_to_100.Label = "100% &transparent"
        Me.grpWCAG_mnu_SetTransparency_to_100.Name = "grpWCAG_mnu_SetTransparency_to_100"
        Me.grpWCAG_mnu_SetTransparency_to_100.OfficeImageId = "BevelShapeGallery"
        Me.grpWCAG_mnu_SetTransparency_to_100.ScreenTip = "100% transparent"
        Me.grpWCAG_mnu_SetTransparency_to_100.ShowImage = True
        Me.grpWCAG_mnu_SetTransparency_to_100.SuperTip = "Will set the image back panel to 100% transparent"
        '
        'grpWCAG_tool_convertAllStyles_toBlack
        '
        Me.grpWCAG_tool_convertAllStyles_toBlack.Label = "Set all text to &black"
        Me.grpWCAG_tool_convertAllStyles_toBlack.Name = "grpWCAG_tool_convertAllStyles_toBlack"
        Me.grpWCAG_tool_convertAllStyles_toBlack.OfficeImageId = "BevelShapeGallery"
        Me.grpWCAG_tool_convertAllStyles_toBlack.ScreenTip = "All text to black"
        Me.grpWCAG_tool_convertAllStyles_toBlack.ShowImage = True
        Me.grpWCAG_tool_convertAllStyles_toBlack.SuperTip = "Will set all text in the current document to black"
        '
        'grpWCAG_tool_tableHeaderColour_all
        '
        Me.grpWCAG_tool_tableHeaderColour_all.Label = "Change Table Header Row Colour"
        Me.grpWCAG_tool_tableHeaderColour_all.Name = "grpWCAG_tool_tableHeaderColour_all"
        Me.grpWCAG_tool_tableHeaderColour_all.OfficeImageId = "ViewBackToColorView"
        Me.grpWCAG_tool_tableHeaderColour_all.ScreenTip = "Table(s) Heder/Row Colour"
        Me.grpWCAG_tool_tableHeaderColour_all.ShowImage = True
        Me.grpWCAG_tool_tableHeaderColour_all.SuperTip = "Will allow you to select a custom colour for the header row of all tables in the " &
    "document"
        '
        'grpReport_PlH_convertToInline_findAllFloatingTables
        '
        Me.grpReport_PlH_convertToInline_findAllFloatingTables.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grpReport_PlH_convertToInline_findAllFloatingTables.KeyTip = "AD"
        Me.grpReport_PlH_convertToInline_findAllFloatingTables.Label = "Placeholder Map"
        Me.grpReport_PlH_convertToInline_findAllFloatingTables.Name = "grpReport_PlH_convertToInline_findAllFloatingTables"
        Me.grpReport_PlH_convertToInline_findAllFloatingTables.OfficeImageId = "_3DSurfaceMaterialClassic"
        Me.grpReport_PlH_convertToInline_findAllFloatingTables.ScreenTip = "Placeholder Map"
        Me.grpReport_PlH_convertToInline_findAllFloatingTables.ShowImage = True
        Me.grpReport_PlH_convertToInline_findAllFloatingTables.SuperTip = resources.GetString("grpReport_PlH_convertToInline_findAllFloatingTables.SuperTip")
        '
        'grpRbn_Mgmnt
        '
        Me.grpRbn_Mgmnt.Items.Add(Me.grpRbn_Mgmnt_mnu_00)
        Me.grpRbn_Mgmnt.Label = "Ribbon Mgmnt"
        Me.grpRbn_Mgmnt.Name = "grpRbn_Mgmnt"
        '
        'grpRbn_Mgmnt_mnu_00
        '
        Me.grpRbn_Mgmnt_mnu_00.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grpRbn_Mgmnt_mnu_00.Items.Add(Me.grpRbn_Mgmnt_removeRbn)
        Me.grpRbn_Mgmnt_mnu_00.Items.Add(Me.Separator54)
        Me.grpRbn_Mgmnt_mnu_00.KeyTip = "RM"
        Me.grpRbn_Mgmnt_mnu_00.Label = "Ribbon Mgmnt"
        Me.grpRbn_Mgmnt_mnu_00.Name = "grpRbn_Mgmnt_mnu_00"
        Me.grpRbn_Mgmnt_mnu_00.OfficeImageId = "MindMapChangeTopic"
        Me.grpRbn_Mgmnt_mnu_00.ShowImage = True
        '
        'grpRbn_Mgmnt_removeRbn
        '
        Me.grpRbn_Mgmnt_removeRbn.Label = "&Remove the Ribbon"
        Me.grpRbn_Mgmnt_removeRbn.Name = "grpRbn_Mgmnt_removeRbn"
        Me.grpRbn_Mgmnt_removeRbn.OfficeImageId = "TraceDependentRemoveArrows"
        Me.grpRbn_Mgmnt_removeRbn.ScreenTip = "Remove the Ribbon"
        Me.grpRbn_Mgmnt_removeRbn.ShowImage = True
        '
        'Separator54
        '
        Me.Separator54.Name = "Separator54"
        '
        'grpTst_LoadFromWeb
        '
        Me.grpTst_LoadFromWeb.Items.Add(Me.grpRbn_Downloads_mnu_00)
        Me.grpTst_LoadFromWeb.Label = "Download support files"
        Me.grpTst_LoadFromWeb.Name = "grpTst_LoadFromWeb"
        '
        'grpRbn_Downloads_mnu_00
        '
        Me.grpRbn_Downloads_mnu_00.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grpRbn_Downloads_mnu_00.Items.Add(Me.grpTst_LoadFromResources_getStylesGuide)
        Me.grpRbn_Downloads_mnu_00.Items.Add(Me.grpTst_LoadFromResources_getStylesGuide_Accessible)
        Me.grpRbn_Downloads_mnu_00.Items.Add(Me.grpTst_LoadFromWeb_getStylesGuide)
        Me.grpRbn_Downloads_mnu_00.Items.Add(Me.grpTst_LoadFromWeb_getStylesGuide_Accessible)
        Me.grpRbn_Downloads_mnu_00.Items.Add(Me.Separator53)
        Me.grpRbn_Downloads_mnu_00.Items.Add(Me.grpTst_LoadFromResources_getTemplate)
        Me.grpRbn_Downloads_mnu_00.Items.Add(Me.grpTst_LoadFromResources_getThemeFile)
        Me.grpRbn_Downloads_mnu_00.Items.Add(Me.Separator56)
        Me.grpRbn_Downloads_mnu_00.Items.Add(Me.grpTst_LoadFromResources_getRptPrtExample)
        Me.grpRbn_Downloads_mnu_00.Items.Add(Me.grpTst_LoadFromResources_getRptLndExample)
        Me.grpRbn_Downloads_mnu_00.Items.Add(Me.grpTst_LoadFromResources_getRptBrfExample)
        Me.grpRbn_Downloads_mnu_00.Items.Add(Me.grpTst_LoadFromWeb_getTemplate)
        Me.grpRbn_Downloads_mnu_00.Label = "Resource Downloads"
        Me.grpRbn_Downloads_mnu_00.Name = "grpRbn_Downloads_mnu_00"
        Me.grpRbn_Downloads_mnu_00.OfficeImageId = "MindMapChangeTopic"
        Me.grpRbn_Downloads_mnu_00.ScreenTip = "Resource Downloads"
        Me.grpRbn_Downloads_mnu_00.ShowImage = True
        Me.grpRbn_Downloads_mnu_00.SuperTip = resources.GetString("grpRbn_Downloads_mnu_00.SuperTip")
        '
        'grpTst_LoadFromResources_getStylesGuide
        '
        Me.grpTst_LoadFromResources_getStylesGuide.Label = "&Styles Guide (from resources)"
        Me.grpTst_LoadFromResources_getStylesGuide.Name = "grpTst_LoadFromResources_getStylesGuide"
        Me.grpTst_LoadFromResources_getStylesGuide.OfficeImageId = "BevelShapeGallery"
        Me.grpTst_LoadFromResources_getStylesGuide.ScreenTip = "Download Styles Guide"
        Me.grpTst_LoadFromResources_getStylesGuide.ShowImage = True
        Me.grpTst_LoadFromResources_getStylesGuide.SuperTip = "When activated this tool will download a 'Styles Guide' document."
        '
        'grpTst_LoadFromResources_getStylesGuide_Accessible
        '
        Me.grpTst_LoadFromResources_getStylesGuide_Accessible.Label = "&Accessible Styles Guide (from resources)"
        Me.grpTst_LoadFromResources_getStylesGuide_Accessible.Name = "grpTst_LoadFromResources_getStylesGuide_Accessible"
        Me.grpTst_LoadFromResources_getStylesGuide_Accessible.OfficeImageId = "BevelShapeGallery"
        Me.grpTst_LoadFromResources_getStylesGuide_Accessible.ScreenTip = "Accessible Styles Guide (from resources)"
        Me.grpTst_LoadFromResources_getStylesGuide_Accessible.ShowImage = True
        '
        'grpTst_LoadFromWeb_getStylesGuide
        '
        Me.grpTst_LoadFromWeb_getStylesGuide.Label = "Styles Guide (from web)"
        Me.grpTst_LoadFromWeb_getStylesGuide.Name = "grpTst_LoadFromWeb_getStylesGuide"
        Me.grpTst_LoadFromWeb_getStylesGuide.OfficeImageId = "BevelShapeGallery"
        Me.grpTst_LoadFromWeb_getStylesGuide.ScreenTip = "Download Styles Guide (web)"
        Me.grpTst_LoadFromWeb_getStylesGuide.ShowImage = True
        Me.grpTst_LoadFromWeb_getStylesGuide.SuperTip = """This is a test tool. When activated it downloads the standard styles guide from " &
    "the AA web site."""
        Me.grpTst_LoadFromWeb_getStylesGuide.Visible = False
        '
        'grpTst_LoadFromWeb_getStylesGuide_Accessible
        '
        Me.grpTst_LoadFromWeb_getStylesGuide_Accessible.Label = "Accessible Styles Guide (from web)"
        Me.grpTst_LoadFromWeb_getStylesGuide_Accessible.Name = "grpTst_LoadFromWeb_getStylesGuide_Accessible"
        Me.grpTst_LoadFromWeb_getStylesGuide_Accessible.OfficeImageId = "BevelShapeGallery"
        Me.grpTst_LoadFromWeb_getStylesGuide_Accessible.ScreenTip = "Download Accessible Styles Guide"
        Me.grpTst_LoadFromWeb_getStylesGuide_Accessible.ShowImage = True
        Me.grpTst_LoadFromWeb_getStylesGuide_Accessible.SuperTip = "Download Accessible Styles Guide"
        Me.grpTst_LoadFromWeb_getStylesGuide_Accessible.Visible = False
        '
        'Separator53
        '
        Me.Separator53.Name = "Separator53"
        '
        'grpTst_LoadFromResources_getTemplate
        '
        Me.grpTst_LoadFromResources_getTemplate.Label = "Template (from resources)"
        Me.grpTst_LoadFromResources_getTemplate.Name = "grpTst_LoadFromResources_getTemplate"
        Me.grpTst_LoadFromResources_getTemplate.OfficeImageId = "BevelShapeGallery"
        Me.grpTst_LoadFromResources_getTemplate.ShowImage = True
        Me.grpTst_LoadFromResources_getTemplate.SuperTip = "Template (from resources)"
        Me.grpTst_LoadFromResources_getTemplate.Visible = False
        '
        'grpTst_LoadFromResources_getThemeFile
        '
        Me.grpTst_LoadFromResources_getThemeFile.Label = "Theme file (from resources"
        Me.grpTst_LoadFromResources_getThemeFile.Name = "grpTst_LoadFromResources_getThemeFile"
        Me.grpTst_LoadFromResources_getThemeFile.ShowImage = True
        Me.grpTst_LoadFromResources_getThemeFile.Visible = False
        '
        'Separator56
        '
        Me.Separator56.Name = "Separator56"
        '
        'grpTst_LoadFromResources_getRptPrtExample
        '
        Me.grpTst_LoadFromResources_getRptPrtExample.Label = "Prt report example (from resources)"
        Me.grpTst_LoadFromResources_getRptPrtExample.Name = "grpTst_LoadFromResources_getRptPrtExample"
        Me.grpTst_LoadFromResources_getRptPrtExample.OfficeImageId = "BevelShapeGallery"
        Me.grpTst_LoadFromResources_getRptPrtExample.ScreenTip = "Styles Guide (from resources)"
        Me.grpTst_LoadFromResources_getRptPrtExample.ShowImage = True
        Me.grpTst_LoadFromResources_getRptPrtExample.Visible = False
        '
        'grpTst_LoadFromResources_getRptLndExample
        '
        Me.grpTst_LoadFromResources_getRptLndExample.Label = "Lnd report example (from resources)"
        Me.grpTst_LoadFromResources_getRptLndExample.Name = "grpTst_LoadFromResources_getRptLndExample"
        Me.grpTst_LoadFromResources_getRptLndExample.OfficeImageId = "BevelShapeGallery"
        Me.grpTst_LoadFromResources_getRptLndExample.ShowImage = True
        Me.grpTst_LoadFromResources_getRptLndExample.Visible = False
        '
        'grpTst_LoadFromResources_getRptBrfExample
        '
        Me.grpTst_LoadFromResources_getRptBrfExample.Label = "Brief report example (from resources)"
        Me.grpTst_LoadFromResources_getRptBrfExample.Name = "grpTst_LoadFromResources_getRptBrfExample"
        Me.grpTst_LoadFromResources_getRptBrfExample.OfficeImageId = "BevelShapeGallery"
        Me.grpTst_LoadFromResources_getRptBrfExample.ShowImage = True
        Me.grpTst_LoadFromResources_getRptBrfExample.Visible = False
        '
        'grpTst_LoadFromWeb_getTemplate
        '
        Me.grpTst_LoadFromWeb_getTemplate.Label = "&Template (from web)"
        Me.grpTst_LoadFromWeb_getTemplate.Name = "grpTst_LoadFromWeb_getTemplate"
        Me.grpTst_LoadFromWeb_getTemplate.OfficeImageId = "BevelShapeGallery"
        Me.grpTst_LoadFromWeb_getTemplate.ScreenTip = "Download Template"
        Me.grpTst_LoadFromWeb_getTemplate.ShowImage = True
        Me.grpTst_LoadFromWeb_getTemplate.SuperTip = """This is a test tool. When activated it downloads the current template from the A" &
    "A web site."""
        Me.grpTst_LoadFromWeb_getTemplate.Visible = False
        '
        'grpMetaData
        '
        Me.grpMetaData.Items.Add(Me.grpMetaData_Remove_FromDoc)
        Me.grpMetaData.Label = "File security"
        Me.grpMetaData.Name = "grpMetaData"
        '
        'grpMetaData_Remove_FromDoc
        '
        Me.grpMetaData_Remove_FromDoc.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grpMetaData_Remove_FromDoc.KeyTip = "FM"
        Me.grpMetaData_Remove_FromDoc.Label = "Remove Meta Data"
        Me.grpMetaData_Remove_FromDoc.Name = "grpMetaData_Remove_FromDoc"
        Me.grpMetaData_Remove_FromDoc.OfficeImageId = "FileDocumentEncrypt"
        Me.grpMetaData_Remove_FromDoc.ScreenTip = "Remove Meta Data"
        Me.grpMetaData_Remove_FromDoc.ShowImage = True
        Me.grpMetaData_Remove_FromDoc.SuperTip = resources.GetString("grpMetaData_Remove_FromDoc.SuperTip")
        '
        'grpTestTools
        '
        Me.grpTestTools.Items.Add(Me.grpTest_pgNum_getTagStyleMap)
        Me.grpTestTools.Label = "User level test tools"
        Me.grpTestTools.Name = "grpTestTools"
        '
        'grpTest_pgNum_getTagStyleMap
        '
        Me.grpTest_pgNum_getTagStyleMap.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grpTest_pgNum_getTagStyleMap.KeyTip = "FT"
        Me.grpTest_pgNum_getTagStyleMap.Label = "Get Tag Style Map"
        Me.grpTest_pgNum_getTagStyleMap.Name = "grpTest_pgNum_getTagStyleMap"
        Me.grpTest_pgNum_getTagStyleMap.OfficeImageId = "BevelShapeGallery"
        Me.grpTest_pgNum_getTagStyleMap.ScreenTip = "Get Tag Style Map"
        Me.grpTest_pgNum_getTagStyleMap.ShowImage = True
        Me.grpTest_pgNum_getTagStyleMap.SuperTip = resources.GetString("grpTest_pgNum_getTagStyleMap.SuperTip")
        '
        'tab_aa_Home
        '
        Me.tab_aa_Home.Groups.Add(Me.grp_AA_ThemeandHome)
        Me.tab_aa_Home.Groups.Add(Me.grp_buildDocuments)
        Me.tab_aa_Home.Groups.Add(Me.grpTest)
        Me.tab_aa_Home.Groups.Add(Me.grp_SwBuild)
        Me.tab_aa_Home.Groups.Add(Me.Group2)
        Me.tab_aa_Home.KeyTip = "JT"
        Me.tab_aa_Home.Label = "AA Home Tab"
        Me.tab_aa_Home.Name = "tab_aa_Home"
        Me.tab_aa_Home.Position = Me.Factory.RibbonPosition.BeforeOfficeId("TabHome")
        '
        'grp_AA_ThemeandHome
        '
        Me.grp_AA_ThemeandHome.DialogLauncher = RibbonDialogLauncherImpl1
        Me.grp_AA_ThemeandHome.Items.Add(Me.Menu1)
        Me.grp_AA_ThemeandHome.Items.Add(Me.tabThms_mnu_resetStyles1)
        Me.grp_AA_ThemeandHome.Items.Add(Me.btn_colorPicker)
        Me.grp_AA_ThemeandHome.Label = "ACIl Allen Theme"
        Me.grp_AA_ThemeandHome.Name = "grp_AA_ThemeandHome"
        '
        'Menu1
        '
        Me.Menu1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu1.Description = "Allows the user to apply standard ACIL Allen Theme"
        Me.Menu1.Items.Add(Me.tabThms_mnu_Set_btn_applyAATheme)
        Me.Menu1.Items.Add(Me.tabThms_mnu_Set_btn__applyAAThemeStylesTemplate)
        Me.Menu1.Items.Add(Me.Separator6)
        Me.Menu1.Items.Add(Me.tabThms_mnu_Set_btn_attachNormalTemplate)
        Me.Menu1.Items.Add(Me.tabThms_mnu_Set_btn_attachAATemplate)
        Me.Menu1.Items.Add(Me.Separator60)
        Me.Menu1.Items.Add(Me.tabThms_mnu_Set_btn_getAttachedTemplate)
        Me.Menu1.Items.Add(Me.tabThms_mnu_Set_btn_ActivateTabPGS)
        Me.Menu1.Items.Add(Me.tabThms_mnu_Set_btn_PGSToggle)
        Me.Menu1.Label = "Set  AA Theme"
        Me.Menu1.Name = "Menu1"
        Me.Menu1.OfficeImageId = "ThemesGallery"
        Me.Menu1.ShowImage = True
        '
        'tabThms_mnu_Set_btn_applyAATheme
        '
        Me.tabThms_mnu_Set_btn_applyAATheme.Label = "Apply the ACIL Allen theme to the current document"
        Me.tabThms_mnu_Set_btn_applyAATheme.Name = "tabThms_mnu_Set_btn_applyAATheme"
        Me.tabThms_mnu_Set_btn_applyAATheme.OfficeImageId = "ThemesGallery"
        Me.tabThms_mnu_Set_btn_applyAATheme.ShowImage = True
        Me.tabThms_mnu_Set_btn_applyAATheme.SuperTip = "The ACIL Allen theme is applied to the current document. Nothing else in that doc" &
    "ument changes."
        '
        'tabThms_mnu_Set_btn__applyAAThemeStylesTemplate
        '
        Me.tabThms_mnu_Set_btn__applyAAThemeStylesTemplate.Label = "Apply ACIL Allen theme, styles and template to the current document"
        Me.tabThms_mnu_Set_btn__applyAAThemeStylesTemplate.Name = "tabThms_mnu_Set_btn__applyAAThemeStylesTemplate"
        Me.tabThms_mnu_Set_btn__applyAAThemeStylesTemplate.OfficeImageId = "ThemesGallery"
        Me.tabThms_mnu_Set_btn__applyAAThemeStylesTemplate.ShowImage = True
        Me.tabThms_mnu_Set_btn__applyAAThemeStylesTemplate.SuperTip = resources.GetString("tabThms_mnu_Set_btn__applyAAThemeStylesTemplate.SuperTip")
        '
        'Separator6
        '
        Me.Separator6.Name = "Separator6"
        '
        'tabThms_mnu_Set_btn_attachNormalTemplate
        '
        Me.tabThms_mnu_Set_btn_attachNormalTemplate.Label = "Attach the Normal template to the document"
        Me.tabThms_mnu_Set_btn_attachNormalTemplate.Name = "tabThms_mnu_Set_btn_attachNormalTemplate"
        Me.tabThms_mnu_Set_btn_attachNormalTemplate.OfficeImageId = "ThemesGallery"
        Me.tabThms_mnu_Set_btn_attachNormalTemplate.ScreenTip = "Adjust theme, styles, then attach template"
        Me.tabThms_mnu_Set_btn_attachNormalTemplate.ShowImage = True
        '
        'tabThms_mnu_Set_btn_attachAATemplate
        '
        Me.tabThms_mnu_Set_btn_attachAATemplate.Label = "Attach the ACIL Allen template to the current document"
        Me.tabThms_mnu_Set_btn_attachAATemplate.Name = "tabThms_mnu_Set_btn_attachAATemplate"
        Me.tabThms_mnu_Set_btn_attachAATemplate.OfficeImageId = "ThemesGallery"
        Me.tabThms_mnu_Set_btn_attachAATemplate.ShowImage = True
        '
        'Separator60
        '
        Me.Separator60.Name = "Separator60"
        '
        'tabThms_mnu_Set_btn_getAttachedTemplate
        '
        Me.tabThms_mnu_Set_btn_getAttachedTemplate.Label = "Get the name of the attached template"
        Me.tabThms_mnu_Set_btn_getAttachedTemplate.Name = "tabThms_mnu_Set_btn_getAttachedTemplate"
        Me.tabThms_mnu_Set_btn_getAttachedTemplate.OfficeImageId = "ThemesGallery"
        Me.tabThms_mnu_Set_btn_getAttachedTemplate.ShowImage = True
        '
        'tabThms_mnu_Set_btn_ActivateTabPGS
        '
        Me.tabThms_mnu_Set_btn_ActivateTabPGS.Label = "Activate Pages and Sections tab"
        Me.tabThms_mnu_Set_btn_ActivateTabPGS.Name = "tabThms_mnu_Set_btn_ActivateTabPGS"
        Me.tabThms_mnu_Set_btn_ActivateTabPGS.OfficeImageId = "ThemesGallery"
        Me.tabThms_mnu_Set_btn_ActivateTabPGS.ShowImage = True
        Me.tabThms_mnu_Set_btn_ActivateTabPGS.Visible = False
        '
        'tabThms_mnu_Set_btn_PGSToggle
        '
        Me.tabThms_mnu_Set_btn_PGSToggle.Label = "Make Pages and Sections tab invisible"
        Me.tabThms_mnu_Set_btn_PGSToggle.Name = "tabThms_mnu_Set_btn_PGSToggle"
        Me.tabThms_mnu_Set_btn_PGSToggle.OfficeImageId = "ThemesGallery"
        Me.tabThms_mnu_Set_btn_PGSToggle.ShowImage = True
        Me.tabThms_mnu_Set_btn_PGSToggle.Visible = False
        '
        'tabThms_mnu_resetStyles1
        '
        Me.tabThms_mnu_resetStyles1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.tabThms_mnu_resetStyles1.Items.Add(Me.tabThms_btn_resetStylesForRptPrt)
        Me.tabThms_mnu_resetStyles1.Items.Add(Me.tabThms_btn_resetStylesForRptLnd)
        Me.tabThms_mnu_resetStyles1.Items.Add(Me.tabThms_btn_resetStylesForRptBrf)
        Me.tabThms_mnu_resetStyles1.Label = "Reset Styles"
        Me.tabThms_mnu_resetStyles1.Name = "tabThms_mnu_resetStyles1"
        Me.tabThms_mnu_resetStyles1.OfficeImageId = "ResetFormatting"
        Me.tabThms_mnu_resetStyles1.ShowImage = True
        Me.tabThms_mnu_resetStyles1.SuperTip = "Will reset the styles for the appropriate report type"
        '
        'tabThms_btn_resetStylesForRptPrt
        '
        Me.tabThms_btn_resetStylesForRptPrt.Label = "Reset for AA Portrait Report"
        Me.tabThms_btn_resetStylesForRptPrt.Name = "tabThms_btn_resetStylesForRptPrt"
        Me.tabThms_btn_resetStylesForRptPrt.OfficeImageId = "PageLayoutTemplatesGallery"
        Me.tabThms_btn_resetStylesForRptPrt.ShowImage = True
        '
        'tabThms_btn_resetStylesForRptLnd
        '
        Me.tabThms_btn_resetStylesForRptLnd.Label = "Reset for AA Landscape Report"
        Me.tabThms_btn_resetStylesForRptLnd.Name = "tabThms_btn_resetStylesForRptLnd"
        Me.tabThms_btn_resetStylesForRptLnd.OfficeImageId = "PageLayouts"
        Me.tabThms_btn_resetStylesForRptLnd.ShowImage = True
        '
        'tabThms_btn_resetStylesForRptBrf
        '
        Me.tabThms_btn_resetStylesForRptBrf.Label = "Reset for AA Brief Report"
        Me.tabThms_btn_resetStylesForRptBrf.Name = "tabThms_btn_resetStylesForRptBrf"
        Me.tabThms_btn_resetStylesForRptBrf.OfficeImageId = "PictureEdgeEffectsGallery"
        Me.tabThms_btn_resetStylesForRptBrf.ShowImage = True
        '
        'btn_colorPicker
        '
        Me.btn_colorPicker.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btn_colorPicker.Label = "Colour Picker"
        Me.btn_colorPicker.Name = "btn_colorPicker"
        Me.btn_colorPicker.OfficeImageId = "BevelShapeGallery"
        Me.btn_colorPicker.ShowImage = True
        Me.btn_colorPicker.Visible = False
        '
        'grp_buildDocuments
        '
        Me.grp_buildDocuments.Items.Add(Me.tbHome_mnu_CreateReport)
        Me.grp_buildDocuments.Items.Add(Me.Separator90)
        Me.grp_buildDocuments.Items.Add(Me.tbHome_grpLetter_standaloneLetter)
        Me.grp_buildDocuments.Items.Add(Me.tbHome_grpLetter_standaloneMemo)
        Me.grp_buildDocuments.Items.Add(Me.tbHome_mnu_contactDetails)
        Me.grp_buildDocuments.Items.Add(Me.Separator91)
        Me.grp_buildDocuments.Items.Add(Me.btn_update_Fields)
        Me.grp_buildDocuments.Items.Add(Me.Separator95)
        Me.grp_buildDocuments.Items.Add(Me.tbHome_btn_ToggleView)
        Me.grp_buildDocuments.Items.Add(Me.Separator94)
        Me.grp_buildDocuments.Items.Add(Me.tbHome_btn_Help)
        Me.grp_buildDocuments.Label = "AA Home Tools"
        Me.grp_buildDocuments.Name = "grp_buildDocuments"
        '
        'tbHome_mnu_CreateReport
        '
        Me.tbHome_mnu_CreateReport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.tbHome_mnu_CreateReport.Items.Add(Me.grpReport_tbHome_btn_buildPortraitReport)
        Me.tbHome_mnu_CreateReport.Items.Add(Me.Separator92)
        Me.tbHome_mnu_CreateReport.Items.Add(Me.grpReport_tbHome_btn_buildLandscapeReport)
        Me.tbHome_mnu_CreateReport.Items.Add(Me.Separator93)
        Me.tbHome_mnu_CreateReport.Items.Add(Me.grpReport_tbHome_btn_buildAABrief)
        Me.tbHome_mnu_CreateReport.Label = "Create Report or Brief"
        Me.tbHome_mnu_CreateReport.Name = "tbHome_mnu_CreateReport"
        Me.tbHome_mnu_CreateReport.OfficeImageId = "AnimationTriggerAddMenu"
        Me.tbHome_mnu_CreateReport.ShowImage = True
        '
        'grpReport_tbHome_btn_buildPortraitReport
        '
        Me.grpReport_tbHome_btn_buildPortraitReport.Label = "New Portrait Report"
        Me.grpReport_tbHome_btn_buildPortraitReport.Name = "grpReport_tbHome_btn_buildPortraitReport"
        Me.grpReport_tbHome_btn_buildPortraitReport.OfficeImageId = "SizeToGridAccess"
        Me.grpReport_tbHome_btn_buildPortraitReport.ShowImage = True
        '
        'Separator92
        '
        Me.Separator92.Name = "Separator92"
        '
        'grpReport_tbHome_btn_buildLandscapeReport
        '
        Me.grpReport_tbHome_btn_buildLandscapeReport.Label = "New Landscape Report"
        Me.grpReport_tbHome_btn_buildLandscapeReport.Name = "grpReport_tbHome_btn_buildLandscapeReport"
        Me.grpReport_tbHome_btn_buildLandscapeReport.OfficeImageId = "SizeToGridAccess"
        Me.grpReport_tbHome_btn_buildLandscapeReport.ShowImage = True
        '
        'Separator93
        '
        Me.Separator93.Name = "Separator93"
        '
        'grpReport_tbHome_btn_buildAABrief
        '
        Me.grpReport_tbHome_btn_buildAABrief.Label = "New ACIL allen Brief"
        Me.grpReport_tbHome_btn_buildAABrief.Name = "grpReport_tbHome_btn_buildAABrief"
        Me.grpReport_tbHome_btn_buildAABrief.OfficeImageId = "SizeToGridAccess"
        Me.grpReport_tbHome_btn_buildAABrief.ShowImage = True
        '
        'Separator90
        '
        Me.Separator90.Name = "Separator90"
        '
        'tbHome_grpLetter_standaloneLetter
        '
        Me.tbHome_grpLetter_standaloneLetter.Label = "Letter"
        Me.tbHome_grpLetter_standaloneLetter.Name = "tbHome_grpLetter_standaloneLetter"
        Me.tbHome_grpLetter_standaloneLetter.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.tbHome_grpLetter_standaloneLetter.ScreenTip = "Stand alone letter"
        Me.tbHome_grpLetter_standaloneLetter.ShowImage = True
        '
        'tbHome_grpLetter_standaloneMemo
        '
        Me.tbHome_grpLetter_standaloneMemo.Label = "Memo"
        Me.tbHome_grpLetter_standaloneMemo.Name = "tbHome_grpLetter_standaloneMemo"
        Me.tbHome_grpLetter_standaloneMemo.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.tbHome_grpLetter_standaloneMemo.ScreenTip = "Stand alone memo"
        Me.tbHome_grpLetter_standaloneMemo.ShowImage = True
        '
        'tbHome_mnu_contactDetails
        '
        Me.tbHome_mnu_contactDetails.Items.Add(Me.tbHome_mnu_contactDetails_letter)
        Me.tbHome_mnu_contactDetails.Items.Add(Me.tbHome_mnu_contactDetails_memo)
        Me.tbHome_mnu_contactDetails.Label = "Contact details"
        Me.tbHome_mnu_contactDetails.Name = "tbHome_mnu_contactDetails"
        Me.tbHome_mnu_contactDetails.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.tbHome_mnu_contactDetails.ScreenTip = "Change letter or memo contact details"
        Me.tbHome_mnu_contactDetails.ShowImage = True
        '
        'tbHome_mnu_contactDetails_letter
        '
        Me.tbHome_mnu_contactDetails_letter.Items.Add(Me.tbHome_mnu_contactDetails_letter_Melbourne)
        Me.tbHome_mnu_contactDetails_letter.Items.Add(Me.tbHome_mnu_contactDetails_letter_Sydney)
        Me.tbHome_mnu_contactDetails_letter.Items.Add(Me.tbHome_mnu_contactDetails_letter_Brisbane)
        Me.tbHome_mnu_contactDetails_letter.Items.Add(Me.tbHome_mnu_contactDetails_letter_Canberra)
        Me.tbHome_mnu_contactDetails_letter.Items.Add(Me.tbHome_mnu_contactDetails_letter_Perth)
        Me.tbHome_mnu_contactDetails_letter.Items.Add(Me.tbHome_mnu_contactDetails_letter_Adelaide)
        Me.tbHome_mnu_contactDetails_letter.Label = "For &Letterhead"
        Me.tbHome_mnu_contactDetails_letter.Name = "tbHome_mnu_contactDetails_letter"
        Me.tbHome_mnu_contactDetails_letter.OfficeImageId = "BevelShapeGallery"
        Me.tbHome_mnu_contactDetails_letter.ShowImage = True
        '
        'tbHome_mnu_contactDetails_letter_Melbourne
        '
        Me.tbHome_mnu_contactDetails_letter_Melbourne.Label = "&Melbourne"
        Me.tbHome_mnu_contactDetails_letter_Melbourne.Name = "tbHome_mnu_contactDetails_letter_Melbourne"
        Me.tbHome_mnu_contactDetails_letter_Melbourne.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.tbHome_mnu_contactDetails_letter_Melbourne.ShowImage = True
        '
        'tbHome_mnu_contactDetails_letter_Sydney
        '
        Me.tbHome_mnu_contactDetails_letter_Sydney.Label = "&Sydney"
        Me.tbHome_mnu_contactDetails_letter_Sydney.Name = "tbHome_mnu_contactDetails_letter_Sydney"
        Me.tbHome_mnu_contactDetails_letter_Sydney.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.tbHome_mnu_contactDetails_letter_Sydney.ShowImage = True
        '
        'tbHome_mnu_contactDetails_letter_Brisbane
        '
        Me.tbHome_mnu_contactDetails_letter_Brisbane.Label = "&Brisbane"
        Me.tbHome_mnu_contactDetails_letter_Brisbane.Name = "tbHome_mnu_contactDetails_letter_Brisbane"
        Me.tbHome_mnu_contactDetails_letter_Brisbane.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.tbHome_mnu_contactDetails_letter_Brisbane.ShowImage = True
        '
        'tbHome_mnu_contactDetails_letter_Canberra
        '
        Me.tbHome_mnu_contactDetails_letter_Canberra.Label = "&Canberra"
        Me.tbHome_mnu_contactDetails_letter_Canberra.Name = "tbHome_mnu_contactDetails_letter_Canberra"
        Me.tbHome_mnu_contactDetails_letter_Canberra.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.tbHome_mnu_contactDetails_letter_Canberra.ShowImage = True
        '
        'tbHome_mnu_contactDetails_letter_Perth
        '
        Me.tbHome_mnu_contactDetails_letter_Perth.Label = "&Perth"
        Me.tbHome_mnu_contactDetails_letter_Perth.Name = "tbHome_mnu_contactDetails_letter_Perth"
        Me.tbHome_mnu_contactDetails_letter_Perth.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.tbHome_mnu_contactDetails_letter_Perth.ShowImage = True
        '
        'tbHome_mnu_contactDetails_letter_Adelaide
        '
        Me.tbHome_mnu_contactDetails_letter_Adelaide.Label = "&Adelaide"
        Me.tbHome_mnu_contactDetails_letter_Adelaide.Name = "tbHome_mnu_contactDetails_letter_Adelaide"
        Me.tbHome_mnu_contactDetails_letter_Adelaide.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.tbHome_mnu_contactDetails_letter_Adelaide.ShowImage = True
        '
        'tbHome_mnu_contactDetails_memo
        '
        Me.tbHome_mnu_contactDetails_memo.Label = "For &Memo"
        Me.tbHome_mnu_contactDetails_memo.Name = "tbHome_mnu_contactDetails_memo"
        Me.tbHome_mnu_contactDetails_memo.OfficeImageId = "BevelShapeGallery"
        Me.tbHome_mnu_contactDetails_memo.ShowImage = True
        Me.tbHome_mnu_contactDetails_memo.Visible = False
        '
        'Separator91
        '
        Me.Separator91.Name = "Separator91"
        '
        'btn_update_Fields
        '
        Me.btn_update_Fields.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btn_update_Fields.Label = "Update Fields"
        Me.btn_update_Fields.Name = "btn_update_Fields"
        Me.btn_update_Fields.OfficeImageId = "FieldValidationMenu"
        Me.btn_update_Fields.ScreenTip = "Update all fields"
        Me.btn_update_Fields.ShowImage = True
        Me.btn_update_Fields.SuperTip = "Will cause all automatic fields to update including TOC and TOF"
        '
        'Separator95
        '
        Me.Separator95.Name = "Separator95"
        '
        'tbHome_btn_ToggleView
        '
        Me.tbHome_btn_ToggleView.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.tbHome_btn_ToggleView.Label = "Toggle View"
        Me.tbHome_btn_ToggleView.Name = "tbHome_btn_ToggleView"
        Me.tbHome_btn_ToggleView.OfficeImageId = "ContentControlBuildingBlockGallery"
        Me.tbHome_btn_ToggleView.ShowImage = True
        '
        'Separator94
        '
        Me.Separator94.Name = "Separator94"
        '
        'tbHome_btn_Help
        '
        Me.tbHome_btn_Help.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.tbHome_btn_Help.Label = "Help"
        Me.tbHome_btn_Help.Name = "tbHome_btn_Help"
        Me.tbHome_btn_Help.OfficeImageId = "Help"
        Me.tbHome_btn_Help.ShowImage = True
        '
        'grpTest
        '
        Me.grpTest.Items.Add(Me.grpTest_btn_cloneDoc)
        Me.grpTest.Items.Add(Me.grpTest_btn_getTimeStamp)
        Me.grpTest.Items.Add(Me.grpSectOptions_sect_InsertSection_InFront)
        Me.grpTest.Items.Add(Me.grpSectOptions_sect_InsertSection_Behind)
        Me.grpTest.Label = "Test Group"
        Me.grpTest.Name = "grpTest"
        Me.grpTest.Visible = False
        '
        'grpTest_btn_cloneDoc
        '
        Me.grpTest_btn_cloneDoc.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.grpTest_btn_cloneDoc.Label = "Clone current document"
        Me.grpTest_btn_cloneDoc.Name = "grpTest_btn_cloneDoc"
        Me.grpTest_btn_cloneDoc.OfficeImageId = "BevelShapeGallery"
        Me.grpTest_btn_cloneDoc.ShowImage = True
        '
        'grpTest_btn_getTimeStamp
        '
        Me.grpTest_btn_getTimeStamp.Label = "Get Time Stamp"
        Me.grpTest_btn_getTimeStamp.Name = "grpTest_btn_getTimeStamp"
        Me.grpTest_btn_getTimeStamp.OfficeImageId = "BevelShapeGallery"
        Me.grpTest_btn_getTimeStamp.ScreenTip = "Get Time Stamp"
        Me.grpTest_btn_getTimeStamp.ShowImage = True
        '
        'grpSectOptions_sect_InsertSection_InFront
        '
        Me.grpSectOptions_sect_InsertSection_InFront.Label = "Insert Section (in front)"
        Me.grpSectOptions_sect_InsertSection_InFront.Name = "grpSectOptions_sect_InsertSection_InFront"
        Me.grpSectOptions_sect_InsertSection_InFront.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_sect_InsertSection_InFront.ShowImage = True
        '
        'grpSectOptions_sect_InsertSection_Behind
        '
        Me.grpSectOptions_sect_InsertSection_Behind.Label = "Insert Section (behind)"
        Me.grpSectOptions_sect_InsertSection_Behind.Name = "grpSectOptions_sect_InsertSection_Behind"
        Me.grpSectOptions_sect_InsertSection_Behind.OfficeImageId = "BevelShapeGallery"
        Me.grpSectOptions_sect_InsertSection_Behind.ShowImage = True
        '
        'grp_SwBuild
        '
        Me.grp_SwBuild.Items.Add(Me.grpRpt_btn_GlossaryAndAbbreviations)
        Me.grp_SwBuild.Items.Add(Me.grpReport_btn_newDivider_Chpt)
        Me.grp_SwBuild.Items.Add(Me.Menu5)
        Me.grp_SwBuild.Items.Add(Me.Menu2df)
        Me.grp_SwBuild.Items.Add(Me.Menu2)
        Me.grp_SwBuild.Items.Add(Me.Menu6)
        Me.grp_SwBuild.Items.Add(Me.Menu4)
        Me.grp_SwBuild.Items.Add(Me.grpLetters_mnu_swBuilds)
        Me.grp_SwBuild.Label = "Software Build Options"
        Me.grp_SwBuild.Name = "grp_SwBuild"
        Me.grp_SwBuild.Visible = False
        '
        'grpRpt_btn_GlossaryAndAbbreviations
        '
        Me.grpRpt_btn_GlossaryAndAbbreviations.KeyTip = "RG"
        Me.grpRpt_btn_GlossaryAndAbbreviations.Label = "Glossary/Abbrev"
        Me.grpRpt_btn_GlossaryAndAbbreviations.Name = "grpRpt_btn_GlossaryAndAbbreviations"
        Me.grpRpt_btn_GlossaryAndAbbreviations.OfficeImageId = "PivotClearCustomOrdering"
        Me.grpRpt_btn_GlossaryAndAbbreviations.ShowImage = True
        '
        'grpReport_btn_newDivider_Chpt
        '
        Me.grpReport_btn_newDivider_Chpt.Image = Global.AA_GeneralReport_Addin.My.Resources.Resources.NewPart_TG
        Me.grpReport_btn_newDivider_Chpt.KeyTip = "RP"
        Me.grpReport_btn_newDivider_Chpt.Label = "Part Divider"
        Me.grpReport_btn_newDivider_Chpt.Name = "grpReport_btn_newDivider_Chpt"
        Me.grpReport_btn_newDivider_Chpt.ShowImage = True
        Me.grpReport_btn_newDivider_Chpt.SuperTip = """Inserts a new Part Divider at the current cursor position. The report body is se" &
    "parated into Parts, where each Part may have any number of Chapters"""
        '
        'Menu5
        '
        Me.Menu5.Items.Add(Me.grpExecSum_ExecSum)
        Me.Menu5.Items.Add(Me.grpExecSum_ExecSum_Grey)
        Me.Menu5.KeyTip = "RE"
        Me.Menu5.Label = "Exec Summary"
        Me.Menu5.Name = "Menu5"
        Me.Menu5.OfficeImageId = "SummarizeSlide"
        Me.Menu5.ShowImage = True
        Me.Menu5.SuperTip = """Inserts a new Excutive Summary"""
        '
        'grpExecSum_ExecSum
        '
        Me.grpExecSum_ExecSum.Label = "&Exec Summary"
        Me.grpExecSum_ExecSum.Name = "grpExecSum_ExecSum"
        Me.grpExecSum_ExecSum.OfficeImageId = "SummarizeSlide"
        Me.grpExecSum_ExecSum.ShowImage = True
        Me.grpExecSum_ExecSum.SuperTip = """Inserts a formatted Executive Summary section container."""
        '
        'grpExecSum_ExecSum_Grey
        '
        Me.grpExecSum_ExecSum_Grey.Label = "Exec Summary (&Grey)"
        Me.grpExecSum_ExecSum_Grey.Name = "grpExecSum_ExecSum_Grey"
        Me.grpExecSum_ExecSum_Grey.OfficeImageId = "SummarizeSlide"
        Me.grpExecSum_ExecSum_Grey.ShowImage = True
        Me.grpExecSum_ExecSum_Grey.SuperTip = """Inserts a formatted Executive Summary section with an image back panel set to rg" &
    "b(200, 200, 200)."""
        '
        'Menu2df
        '
        Me.Menu2df.Items.Add(Me.tabPgs_grpRpt_btn_buildPrtReport_sw)
        Me.Menu2df.Items.Add(Me.Separator62)
        Me.Menu2df.Items.Add(Me.tabPgs_grpRpt_btn_buildLndReport_sw)
        Me.Menu2df.Items.Add(Me.Separator63)
        Me.Menu2df.Items.Add(Me.tabPgs_grpRpt_btn_buildBrfReport_sw)
        Me.Menu2df.KeyTip = "RN"
        Me.Menu2df.Label = "Create Report or Brief"
        Me.Menu2df.Name = "Menu2df"
        Me.Menu2df.OfficeImageId = "AnimationTriggerAddMenu"
        Me.Menu2df.ShowImage = True
        '
        'tabPgs_grpRpt_btn_buildPrtReport_sw
        '
        Me.tabPgs_grpRpt_btn_buildPrtReport_sw.Label = "Create a new Portrait Report (sw)"
        Me.tabPgs_grpRpt_btn_buildPrtReport_sw.Name = "tabPgs_grpRpt_btn_buildPrtReport_sw"
        Me.tabPgs_grpRpt_btn_buildPrtReport_sw.OfficeImageId = "SizeToGridAccess"
        Me.tabPgs_grpRpt_btn_buildPrtReport_sw.ShowImage = True
        Me.tabPgs_grpRpt_btn_buildPrtReport_sw.SuperTip = "Will create a new standard ACIL Allen portrait report skeleton (all software)"
        '
        'Separator62
        '
        Me.Separator62.Name = "Separator62"
        '
        'tabPgs_grpRpt_btn_buildLndReport_sw
        '
        Me.tabPgs_grpRpt_btn_buildLndReport_sw.Label = "Create a new Landscape Report (sw)"
        Me.tabPgs_grpRpt_btn_buildLndReport_sw.Name = "tabPgs_grpRpt_btn_buildLndReport_sw"
        Me.tabPgs_grpRpt_btn_buildLndReport_sw.OfficeImageId = "SizeToGridAccess"
        Me.tabPgs_grpRpt_btn_buildLndReport_sw.ShowImage = True
        Me.tabPgs_grpRpt_btn_buildLndReport_sw.SuperTip = "Will create a new standard ACIL Allen landscape report skeleton (all software)"
        '
        'Separator63
        '
        Me.Separator63.Name = "Separator63"
        '
        'tabPgs_grpRpt_btn_buildBrfReport_sw
        '
        Me.tabPgs_grpRpt_btn_buildBrfReport_sw.Label = "Create a new ACIL Allen Brief (sw)"
        Me.tabPgs_grpRpt_btn_buildBrfReport_sw.Name = "tabPgs_grpRpt_btn_buildBrfReport_sw"
        Me.tabPgs_grpRpt_btn_buildBrfReport_sw.OfficeImageId = "SizeToGridAccess"
        Me.tabPgs_grpRpt_btn_buildBrfReport_sw.ShowImage = True
        '
        'Menu2
        '
        Me.Menu2.Items.Add(Me.grpRpt_mnu_btn_NewChapter_inFront)
        Me.Menu2.Items.Add(Me.grpRpt_mnu_btn_NewChapter_behind)
        Me.Menu2.KeyTip = "RC"
        Me.Menu2.Label = "New Chapter"
        Me.Menu2.Name = "Menu2"
        Me.Menu2.OfficeImageId = "CompareAndCombine"
        Me.Menu2.ScreenTip = "Inserts a new Chapter"
        Me.Menu2.ShowImage = True
        '
        'grpRpt_mnu_btn_NewChapter_inFront
        '
        Me.grpRpt_mnu_btn_NewChapter_inFront.Label = "New Chapter (in &Front)"
        Me.grpRpt_mnu_btn_NewChapter_inFront.Name = "grpRpt_mnu_btn_NewChapter_inFront"
        Me.grpRpt_mnu_btn_NewChapter_inFront.OfficeImageId = "BevelShapeGallery"
        Me.grpRpt_mnu_btn_NewChapter_inFront.ShowImage = True
        '
        'grpRpt_mnu_btn_NewChapter_behind
        '
        Me.grpRpt_mnu_btn_NewChapter_behind.Label = "New Chapter (&Behind)"
        Me.grpRpt_mnu_btn_NewChapter_behind.Name = "grpRpt_mnu_btn_NewChapter_behind"
        Me.grpRpt_mnu_btn_NewChapter_behind.OfficeImageId = "BevelShapeGallery"
        Me.grpRpt_mnu_btn_NewChapter_behind.ShowImage = True
        Me.grpRpt_mnu_btn_NewChapter_behind.SuperTip = """Inserts a new chapter section behind the section that contains the current curso" &
    "r position."""
        '
        'Menu6
        '
        Me.Menu6.Items.Add(Me.grpOther_bibliography)
        Me.Menu6.Items.Add(Me.grpOther_references)
        Me.Menu6.Items.Add(Me.grpOther_worksCited)
        Me.Menu6.KeyTip = "RB"
        Me.Menu6.Label = "Bibliography"
        Me.Menu6.Name = "Menu6"
        Me.Menu6.OfficeImageId = "BibliographyGallery"
        Me.Menu6.ShowImage = True
        '
        'grpOther_bibliography
        '
        Me.grpOther_bibliography.Label = "&Bibliography"
        Me.grpOther_bibliography.Name = "grpOther_bibliography"
        Me.grpOther_bibliography.OfficeImageId = "ReplaceWithAutoText"
        Me.grpOther_bibliography.ShowImage = True
        '
        'grpOther_references
        '
        Me.grpOther_references.Label = "&References"
        Me.grpOther_references.Name = "grpOther_references"
        Me.grpOther_references.OfficeImageId = "ReplaceWithAutoText"
        Me.grpOther_references.ShowImage = True
        '
        'grpOther_worksCited
        '
        Me.grpOther_worksCited.Label = "&Works Cited"
        Me.grpOther_worksCited.Name = "grpOther_worksCited"
        Me.grpOther_worksCited.OfficeImageId = "ReplaceWithAutoText"
        Me.grpOther_worksCited.ShowImage = True
        '
        'Menu4
        '
        Me.Menu4.Items.Add(Me.grpAppendix_newAppChapter_inFront)
        Me.Menu4.Items.Add(Me.grpAppendix_newAppChapter_behind)
        Me.Menu4.KeyTip = "AC"
        Me.Menu4.Label = "New Appendix/Att"
        Me.Menu4.Name = "Menu4"
        Me.Menu4.OfficeImageId = "TextBoxInsert"
        Me.Menu4.ScreenTip = "New Appendix/Att"
        Me.Menu4.ShowImage = True
        Me.Menu4.SuperTip = " ""Inserts a new page with Appendix/Attachment heading separated by a section brea" &
    "k. Page number and field in footer are linked to Appendix heading."""
        '
        'grpAppendix_newAppChapter_inFront
        '
        Me.grpAppendix_newAppChapter_inFront.Label = "New App/Att (in &Front)"
        Me.grpAppendix_newAppChapter_inFront.Name = "grpAppendix_newAppChapter_inFront"
        Me.grpAppendix_newAppChapter_inFront.OfficeImageId = "TextBoxInsert"
        Me.grpAppendix_newAppChapter_inFront.ScreenTip = "New App/Att (in Front)"
        Me.grpAppendix_newAppChapter_inFront.ShowImage = True
        Me.grpAppendix_newAppChapter_inFront.SuperTip = """Inserts a new App/Att section in front of the Section that contains the current " &
    "cursor position."""
        '
        'grpAppendix_newAppChapter_behind
        '
        Me.grpAppendix_newAppChapter_behind.Label = "New App/Att (&Behind)"
        Me.grpAppendix_newAppChapter_behind.Name = "grpAppendix_newAppChapter_behind"
        Me.grpAppendix_newAppChapter_behind.OfficeImageId = "TextBoxInsert"
        Me.grpAppendix_newAppChapter_behind.ScreenTip = "New App/Att (Behind)"
        Me.grpAppendix_newAppChapter_behind.ShowImage = True
        Me.grpAppendix_newAppChapter_behind.SuperTip = """Inserts a new App/Attsection behind the Section that contains the current cursor" &
    " position."""
        '
        'grpLetters_mnu_swBuilds
        '
        Me.grpLetters_mnu_swBuilds.Items.Add(Me.grpLetter_insertLetter_swBuild)
        Me.grpLetters_mnu_swBuilds.Items.Add(Me.grpLetter_insertMemo_swBuild)
        Me.grpLetters_mnu_swBuilds.Label = "Stationery Software Build"
        Me.grpLetters_mnu_swBuilds.Name = "grpLetters_mnu_swBuilds"
        Me.grpLetters_mnu_swBuilds.OfficeImageId = "MailMergeAddressBlockInsert"
        Me.grpLetters_mnu_swBuilds.ShowImage = True
        '
        'grpLetter_insertLetter_swBuild
        '
        Me.grpLetter_insertLetter_swBuild.Label = "Letter (sw build)"
        Me.grpLetter_insertLetter_swBuild.Name = "grpLetter_insertLetter_swBuild"
        Me.grpLetter_insertLetter_swBuild.OfficeImageId = "BevelShapeGallery"
        Me.grpLetter_insertLetter_swBuild.ShowImage = True
        '
        'grpLetter_insertMemo_swBuild
        '
        Me.grpLetter_insertMemo_swBuild.Label = "memo (sw build)"
        Me.grpLetter_insertMemo_swBuild.Name = "grpLetter_insertMemo_swBuild"
        Me.grpLetter_insertMemo_swBuild.OfficeImageId = "BevelShapeGallery"
        Me.grpLetter_insertMemo_swBuild.ShowImage = True
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.btn_ApplyStdTheme_Manually)
        Me.Group2.Items.Add(Me.mnu_makeStyles)
        Me.Group2.Label = "Group2"
        Me.Group2.Name = "Group2"
        Me.Group2.Visible = False
        '
        'btn_ApplyStdTheme_Manually
        '
        Me.btn_ApplyStdTheme_Manually.Label = "Apply std theme manually"
        Me.btn_ApplyStdTheme_Manually.Name = "btn_ApplyStdTheme_Manually"
        Me.btn_ApplyStdTheme_Manually.OfficeImageId = "BevelShapeGallery"
        Me.btn_ApplyStdTheme_Manually.ShowImage = True
        '
        'mnu_makeStyles
        '
        Me.mnu_makeStyles.Items.Add(Me.btn_styles_makeTableText)
        Me.mnu_makeStyles.Label = "Make Styles"
        Me.mnu_makeStyles.Name = "mnu_makeStyles"
        Me.mnu_makeStyles.OfficeImageId = "BevelShapeGallery"
        Me.mnu_makeStyles.ShowImage = True
        '
        'btn_styles_makeTableText
        '
        Me.btn_styles_makeTableText.Label = "Table text"
        Me.btn_styles_makeTableText.Name = "btn_styles_makeTableText"
        Me.btn_styles_makeTableText.OfficeImageId = "BevelShapeGallery"
        Me.btn_styles_makeTableText.ShowImage = True
        '
        'rbn_aa_Addin00
        '
        Me.Name = "rbn_aa_Addin00"
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.tab_aa_Styles)
        Me.Tabs.Add(Me.tab_aa_Placeholders)
        Me.Tabs.Add(Me.tab_aa_PagesAndSections)
        Me.Tabs.Add(Me.tab_aa_Finalise)
        Me.Tabs.Add(Me.tab_aa_Home)
        Me.tab_aa_Styles.ResumeLayout(False)
        Me.tab_aa_Styles.PerformLayout()
        Me.grp_Styles_AAThemes.ResumeLayout(False)
        Me.grp_Styles_AAThemes.PerformLayout()
        Me.grpStyles_CoverPage.ResumeLayout(False)
        Me.grpStyles_CoverPage.PerformLayout()
        Me.grpStyles_Report.ResumeLayout(False)
        Me.grpStyles_Report.PerformLayout()
        Me.grpStyles_NoNum.ResumeLayout(False)
        Me.grpStyles_NoNum.PerformLayout()
        Me.grpStyles_Appendices.ResumeLayout(False)
        Me.grpStyles_Appendices.PerformLayout()
        Me.grpStyles_Text.ResumeLayout(False)
        Me.grpStyles_Text.PerformLayout()
        Me.grpStyles_Lists.ResumeLayout(False)
        Me.grpStyles_Lists.PerformLayout()
        Me.grpStyles_Emphasis.ResumeLayout(False)
        Me.grpStyles_Emphasis.PerformLayout()
        Me.grpStyles_resetStyles.ResumeLayout(False)
        Me.grpStyles_resetStyles.PerformLayout()
        Me.grpStyles_resetCaptions.ResumeLayout(False)
        Me.grpStyles_resetCaptions.PerformLayout()
        Me.tab_aa_Placeholders.ResumeLayout(False)
        Me.tab_aa_Placeholders.PerformLayout()
        Me.grp_PlaceHolders.ResumeLayout(False)
        Me.grp_PlaceHolders.PerformLayout()
        Me.grp_special_AATableFormatting.ResumeLayout(False)
        Me.grp_special_AATableFormatting.PerformLayout()
        Me.grp_floatingPlaceholders.ResumeLayout(False)
        Me.grp_floatingPlaceholders.PerformLayout()
        Me.grp_Plh_miscPlaceholders.ResumeLayout(False)
        Me.grp_Plh_miscPlaceholders.PerformLayout()
        Me.tab_aa_PagesAndSections.ResumeLayout(False)
        Me.tab_aa_PagesAndSections.PerformLayout()
        Me.grpRpt_CoversAndTOC.ResumeLayout(False)
        Me.grpRpt_CoversAndTOC.PerformLayout()
        Me.grpRpt_ImagePanels.ResumeLayout(False)
        Me.grpRpt_ImagePanels.PerformLayout()
        Me.grpRpt_Report.ResumeLayout(False)
        Me.grpRpt_Report.PerformLayout()
        Me.grpRpt_Appendix.ResumeLayout(False)
        Me.grpRpt_Appendix.PerformLayout()
        Me.grpRpt_sectOptions.ResumeLayout(False)
        Me.grpRpt_sectOptions.PerformLayout()
        Me.grpRpt_CoveringLetter.ResumeLayout(False)
        Me.grpRpt_CoveringLetter.PerformLayout()
        Me.grp_WhatsNew.ResumeLayout(False)
        Me.grp_WhatsNew.PerformLayout()
        Me.grp_Fixes.ResumeLayout(False)
        Me.grp_Fixes.PerformLayout()
        Me.tab_aa_Finalise.ResumeLayout(False)
        Me.tab_aa_Finalise.PerformLayout()
        Me.grp_WaterMarks.ResumeLayout(False)
        Me.grp_WaterMarks.PerformLayout()
        Me.grp_PgNumMgmnt.ResumeLayout(False)
        Me.grp_PgNumMgmnt.PerformLayout()
        Me.grp_Finalise.ResumeLayout(False)
        Me.grp_Finalise.PerformLayout()
        Me.grpWCAG.ResumeLayout(False)
        Me.grpWCAG.PerformLayout()
        Me.grpRbn_Mgmnt.ResumeLayout(False)
        Me.grpRbn_Mgmnt.PerformLayout()
        Me.grpTst_LoadFromWeb.ResumeLayout(False)
        Me.grpTst_LoadFromWeb.PerformLayout()
        Me.grpMetaData.ResumeLayout(False)
        Me.grpMetaData.PerformLayout()
        Me.grpTestTools.ResumeLayout(False)
        Me.grpTestTools.PerformLayout()
        Me.tab_aa_Home.ResumeLayout(False)
        Me.tab_aa_Home.PerformLayout()
        Me.grp_AA_ThemeandHome.ResumeLayout(False)
        Me.grp_AA_ThemeandHome.PerformLayout()
        Me.grp_buildDocuments.ResumeLayout(False)
        Me.grp_buildDocuments.PerformLayout()
        Me.grpTest.ResumeLayout(False)
        Me.grpTest.PerformLayout()
        Me.grp_SwBuild.ResumeLayout(False)
        Me.grp_SwBuild.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents tab_aa_Styles As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents grp_Styles_AAThemes As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents mnu_SetTheme As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents xbtn_mnuThemes_ActivateTabPGS As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tab_aa_Placeholders As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents tab_aa_PagesAndSections As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents tab_aa_Finalise As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents xbtn_mnuThemes_PGSToggle As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents xbtn__mnuThemes_set_AATheme As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents xbtn__mnuThemes_set_AAThemeAndStyles As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpRpt_CoversAndTOC As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpRpt_ImagePanels As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpRpt_Report As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpRpt_btn_GlossaryAndAbbreviations_bblk As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator2 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpRpt_mnu_CreateRpt As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpReport_btn_buildPortraitReport As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator3 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpReport_btn_buildAABrief As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator4 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpReprt_btn_buildLandscapeReport As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpRpt_mnu_NewChapter As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpRpt_mnu_CreateExecSummary As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpReport_btn_ToggleView As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpRpt_mnu_RefreshDocument As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpRpt_mnu_ApplyColour As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpRpt_mnu_btn_NewChapter_inFront_bblk As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpRpt_mnu_btn_NewChapter_behind_bblk As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpExecSum_ExecSum_bblk As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpExecSum_ExecSum_Grey_bblk As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_PlaceHolders As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpPlh_btn_buildCustomTable As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStyles_CoverPage As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpStylesES_StyleSet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesES_Heading2_ES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesES_Heading3_ES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesES_Heading4_ES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesES_Heading5_ES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStyles_Report As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpStyles_NoNum As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpStyles_Appendices As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpStyles_Text As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpStyles_Lists As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpStyles_Emphasis As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpStyles_resetStyles As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpStyles_resetCaptions As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpStylesRpt_StyleSet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesRpt_Heading2_Rpt As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesRpt_Heading3_Rpt As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesRpt_Heading4_Rpt As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesRpt_Heading5_Rpt As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator5 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpStyles_mnu_Heading3Numbering As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpStylesRpt_HeadingNoNum_StyleSet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesRpt_Heading2NoNum_Rpt As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesRpt_Heading3NoNum_Rpt As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesRpt_Heading4NoNum_Rpt As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesRpt_Heading5NoNum_Rpt As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStyles_mnu_Heading3Numbering_btn_on As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStyles_mnu_Heading3Numbering_btn_off As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents gal_CoverPages As Microsoft.Office.Tools.Ribbon.RibbonGallery
    Friend WithEvents gal_CoverPages_btn_deleteCoverPage As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpRpt_Appendix As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpRpt_sectOptions As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents mnuCloseDocuments000 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpRpt_sectOptions_btn_delSection As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tab_aa_Home As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents grp_AA_ThemeandHome As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Menu1 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents tabThms_mnu_Set_btn_applyAATheme As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tabThms_mnu_Set_btn__applyAAThemeStylesTemplate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator6 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents tabThms_mnu_Set_btn_ActivateTabPGS As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tabThms_mnu_Set_btn_PGSToggle As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tabThms_mnu_Set_btn_attachNormalTemplate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_mnu01_SelectedText As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_mnu01_SelectedTblCells As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_mnu01_ImageBackPanel As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator7 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpReport_mnu_CaseStudies_RecolourLogo_toWhite As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_mnu_CaseStudies_RecolourLogo_Reset As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator8 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpReport_mnu_CaseStudies_RecolourFooter_toWhite As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_mnu_CaseStudies_RecolourFooter_Reset As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpViewTools_Refresh_Stationery_Ref As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator9 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpViewTools_Refresh_mnu_TOC As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpViewTools_Refresh_mnu_Chapters As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpViewTools_Refresh_mnu_Parts As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator10 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpViewTools_Refresh_mnu_Tables As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpViewTools_Refresh_mnu_Figures As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpViewTools_Refresh_mnu_Boxes As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator11 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpViewTools_Refresh_mnu_All As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator12 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpViewTools_Refresh_mnu_Every As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents mnu_grpViewTools_Refresh_btn_setRefFldNotBold As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator13 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpSectOptions_submnu_LndWidthOptions As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpSectOptions_sect_InsertSectionBounded_Lnd As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpSectOptions_sect_InsertSectionBounded_Lnd_wide As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu3 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpSectOptions_sect_InsertSectionBounded_Prt As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpSectOptions_sect_InsertSectionBounded_Prt_wide As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator14 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpSectOptions_sect_InsertSection_AtSelection As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpOther_mnuHFS As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpSectOptions_header_ClearTextandShapes As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpSectOptions_footer_ClearText As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpSectOptions_footer_ClearTextandPageNum As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator15 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpSectOptions_footer_clearSubTitleField As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator16 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpSectOptions_footer_resetText As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator17 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpOther_mnuHFS_sub00_restoreHF As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpSectOptions_hfs_restoreHF_ES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpSectOptions_hfs_restoreHF_RP As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpSectOptions_hfs_restoreHF_AP As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpSectOptions_mnu_ResetLndPrt As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpSectOptions_mnu_sub00_ResetLndPrt_Lnd As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpSectOptions_resetTo_Lnd_ES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpSectOptions_resetTo_Lnd_RP As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpSectOptions_resetTo_Lnd_AP As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpSectOptions_mnu_sub01_ResetLndPrt_Ln As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpSectOptions_resetTo_Prt_ES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpSectOptions_resetTo_Prt_RP As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpSectOptions_resetTo_Prt_AP As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpSectOptions_mnu_ResetResizeLandscape As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpSectOptions_resizeTo_Landscape As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpSectOptions_resizeTo_Portrait As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator18 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpSectOptions_resize_toggleWidth As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents mnu_grpReport_Columns As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpReport_Columns_04 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_Columns_03 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_Columns_02 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_Columns_02_LeftWide As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_Columns_02_RightWide As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_Columns_01 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator19 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpAppendix_mnu01 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpReport_mnu_NewAppAtt As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpAppendix_newAppPart As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpAppendix_newAttPart As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpAppendix_newAppChapter_inFront_bblk As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpAppendix_newAppChapter_behind_bblk As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpCntsPages As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpCoversToc_mnu2 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpToc_TOC_insertSection As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator20 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpToc_TOC_insertLevels_1_to_1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpToc_TOC_insertLevels_1_to_2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpToc_TOC_insertLevels_1_to_3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator21 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpToc_TOC_update As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpContactsPages_FrontPage_AckOfCountry As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpContactsPages_FrontPage As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator22 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpContactsPages_BackPage As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator23 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpContactsPages_mnu_2 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Separator24 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpContactsPages_CopyrightStatement As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpContactsPages_Disclaimer As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpContactsPages_ReportTo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpContactsPages_ProposalTo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpCoversToc_mnu_Images As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpImageHandling_mnu_ImgSection As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpImageHandling_mnu_FillBackPanel As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpCpImages_ImageFromFile As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpCpImages_ImageFromClip As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator26 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpCpImages_BackPanelFill_RawImageFromFile As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator25 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpCpImages_Delete_SmallPict As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpImageHandling_insert_BackPanel As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator27 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpImageHandling_delete_BackPanel As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpImageHandling_BackPanelFill_FromFile As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpImageHandling_BackPanelFill_FromClipBoard As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator28 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpImageHandling_BackPanelFill_RawImageFromFile As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator29 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpImageHandling_Reset_backcolour As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpImageHandling_Reset_backcolour_to_CaseStudyGrey As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpImageHandling_Custom_backcolour As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator30 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpImageHandling_submnu_FillBackPanel_SetTransparency As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents submnu_SetTransparency_to_0 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents submnu_SetTransparency_to_25 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents submnu_SetTransparency_to_50 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents submnu_SetTransparency_to_75 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents submnu_SetTransparency_to_100 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents mnu_SetBackPanel_to_BannerHeight As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents mnuCloseDocuments161 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpBoxes_Box As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_AppendixBox As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_ESBox As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_LTBox As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator31 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpBoxes_CaptionAndHeading As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_CaptionAndHeadingES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_CaptionAndHeadingApp As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents mnuCloseDocuments2233 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents mnuCloseDocuments1 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpBoxes_BoxTextBoldItalic As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_BoxText As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_SideHeading1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_SideHeading2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator32 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpBoxes_BoxListBullet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_BoxListBullet2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_BoxListBullet3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator33 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpBoxes_BoxListNumber As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_BoxListNumber2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_BoxListNumber3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator34 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpBoxes_BoxQuote As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_BoxQuoteListBullet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_BoxQuoteSource As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator35 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpBoxes_boxContent_mnu As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpBoxes_deleteBoxContent As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_fillWithExampleStyles As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_ToES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_ToBox1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_ToApp As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator36 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpBoxes_ToLT As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents mnu_grpBoxes_Recommendations As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpBoxes_Recommendation As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_RecommendationES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents mnu_grpBoxes_Findings As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpPullouts_mnu01 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpReport_mnu_CaseStudies As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Separator37 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents mnuCloseDocuments16 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents mnuCloseDocuments33 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpStylesRpt_mnu_tbls_00 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Separator38 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpBoxes_KeyFinding As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_KeyFindingES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpPullouts_emphasisBox_Left As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpPullouts_emphasisBox_Centre As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpPullouts_emphasisBox_Right As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_mnu_CaseStudies_FullPage As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_mnu_CaseStudies_HalfPage As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator39 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpReport_mnu_CaseStudies_CaseStudyHeading As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpFigures_Figure As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator44 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpFigures_Appendix As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator43 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpFigures_ES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator42 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpFigures_LT As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator41 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpFigures_CaptionAndHeading As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpFigures_CaptionAndHeadingApp As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpFigures_CaptionAndHeadingES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator40 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpFigures_StyleForSubHeadings As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpFigures_convertToES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpFigures_convertToStd As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpFigures_convertToApp As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator45 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpFigures_convertToLT As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_fillCellsWithCustomColour As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator46 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpTbls_setTableTextCustomColour As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbl_Styles As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpTbls_TableTextStyle As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_TableColumnHeadingsStyle As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_TableUnitsRowStyle As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_TableListBullet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_TableListBullet2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_TableListBullet3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_ListNumber As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_ListNumber2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_ListNumber3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_TableSideHeading1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_TableSideHeading2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_Quote As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_QuoteListBullet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_QuoteSource As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbl_Styles_ExampleStyleSets As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpTbls_StyleSet_TableQuote As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_StyleSet_TableListBullets As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_StyleSet_TableListNumbers As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_ColourCells As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_ColourHeadingsRow As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_ColourUnitsRow As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_AllBorders As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_AllBordersRemove As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_Plh_mnu_TableListBulletsStyles As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpTbls_Plh_mnu_TableListNumberingStyles As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpTbls_Plh_mnu_SideHeadingStyles As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpTbls_Plh_mnu_QuoteStyles As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Separator47 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Separator48 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Separator49 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Separator50 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Separator51 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents mnuCloseDocuments4 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpTbls_convertTabletoES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_convertTabletoStd As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_convertTabletoApp As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator52 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpTbls_convertTabletoLT As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_AllStyles_small As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpTbls_TableTextStyle_small As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_WaterMarks As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grp_PgNumMgmnt As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grp_Finalise As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpWCAG As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpTst_LoadFromWeb As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpMetaData As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpTestTools As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpTest_pgNum_getTagStyleMap As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpMetaData_Remove_FromDoc As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpRbn_Downloads_mnu_00 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpTst_LoadFromWeb_getTemplate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTst_LoadFromWeb_getStylesGuide As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTst_LoadFromWeb_getStylesGuide_Accessible As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTst_LoadFromResources_getTemplate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTst_LoadFromResources_getRptPrtExample As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTst_LoadFromResources_getStylesGuide_Accessible As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator53 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpRbn_Mgmnt As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpRbn_Mgmnt_mnu_00 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpRbn_Mgmnt_removeRbn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator54 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpTbls_TableColumnHeadingsStyle_small As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_TableUnitsRowStyle_small As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_Plh_mnu_TableListBulletsStyles_small As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpTbls_TableListBullet_small As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_TableListBullet2_small As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_TableListBullet3_small As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_Plh_mnu_TableListNumberingStyles_small As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpTbls_ListNumber_small As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_ListNumber2_small As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator57 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpTbls_ListNumber3_small As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator58 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpPlh_mnu_TblSideHeadings_small As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpTbls_TableSideHeading1_small As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_TableSideHeading2_small As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpPlh_mnu_TblQuoteStyles_small As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpTbls_Quote_small As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_QuoteListBullet_small As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_QuoteSource_small As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator59 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpTbl_Styles_ExampleStyleSets_Small As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpTbls_StyleSet_TableQuote_small As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_StyleSet_TableListBullets_small As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbls_StyleSet_TableListNumbers_small As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpPlh_mnu_TblPlaceholders As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpPlh_mnu_SourceAndNote As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpPlh_mnu_DeleteTable As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpTest As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpTest_btn_cloneDoc As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesTools_to_PrintDefault As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesTools_to_DisplayDefault As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesTools_resetCaptions As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTst_LoadFromResources_getThemeFile As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator56 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpTst_LoadFromResources_getRptLndExample As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTst_LoadFromResources_getRptBrfExample As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tabThms_btn_resetStylesForRptPrt As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tabThms_btn_resetStylesForRptLnd As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tabThms_btn_resetStylesForRptBrf As Microsoft.Office.Tools.Ribbon.RibbonButton
    Public WithEvents tabThms_mnu_resetStyles1 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents tabThms_mnu_Set_btn_attachAATemplate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator60 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents tabThms_mnu_Set_btn_getAttachedTemplate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tbStyles_grpResetStyles_mnu_ResetStyles As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents tabStyles_btn_resetStylesForRptPrt As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tabStyles_btn_resetStylesForRptLnd As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tabStyles_btn_resetStylesForRptBrf As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesApp_StyleSet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesApp_Heading1_App As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesApp_Heading2_App As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesApp_Heading3_App As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesApp_Heading4_App As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesApp_Heading5_App As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTest_btn_getTimeStamp As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpSectOptions_sect_InsertSection_InFront As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpSectOptions_sect_InsertSection_Behind As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpRpt_CoveringLetter As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpLetter_insertLetter As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpLetter_insertMemo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpCoveringLetter_mnu6 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents mnuCloseDocuments777 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Separator61 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpLetter_btn_forMemo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_WhatsNew As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpWhatsNew_Form As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_SwBuild As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpLetters_mnu_swBuilds As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpLetter_insertLetter_swBuild As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpLetter_insertMemo_swBuild As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents mnuCloseDocuments11 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpLetter_LtrHead1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpLetter_LtrHead2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpLetter_LtrHead3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpLetter_delReport As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu2df As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents tabPgs_grpRpt_btn_buildPrtReport_sw As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator62 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents tabPgs_grpRpt_btn_buildLndReport_sw As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator63 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents tabPgs_grpRpt_btn_buildBrfReport_sw As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpLetter_Melbourne As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpLetter_Sydney As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpLetter_Brisbane As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpLetter_Canberra As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpLetter_Perth As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpLetter_Adelaide As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_btn_newDivider_Chpt_bblk As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_btn_newDivider_Chpt As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu2 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpRpt_mnu_btn_NewChapter_inFront As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpRpt_mnu_btn_NewChapter_behind As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu4 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpAppendix_newAppChapter_inFront As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpAppendix_newAppChapter_behind As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu5 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpExecSum_ExecSum As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpExecSum_ExecSum_Grey As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpRpt_btn_GlossaryAndAbbreviations As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpRpt_mnu_Bibliography As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpOther_bibliography_bblk As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpOther_references_bblk As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpOther_worksCited_bblk As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu6 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpOther_bibliography As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpOther_references As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpOther_worksCited As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btn_colorPicker As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btn_ApplyStdTheme_Manually As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesText_BodyText As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesRpt_Intro As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesOther_Quote As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesOther_QuoteBlt As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesOther_QuoteSource As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesLists_List1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesLists_List2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesLists_List3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator55 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpStylesLists_ListNumber1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesLists_ListNumber2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStylesLists_ListNumber3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tbStyles_mnu_Emphasis As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpPullouts_emphasisBox_TextStyle_Left_2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpPullouts_emphasisBox_TextStyle_Centre_2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpPullouts_emphasisBox_TextStyle_Right_2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTblsPlh_HeadingAndSource As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTblsPlh_HeadingAndSourceApp As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTblsPlh_HeadingAndSourceES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTblsPlh_CaptionAndHeading As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTblsPlh_CaptionAndHeadingApp As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTblsPlh_CaptionAndHeadingES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTblsPlh_AddTable_Simple As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator64 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Separator65 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpTblsPlh_SourceLabelAndStyle As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTblsPlh_NoteLabelAndStyle As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTblsPlh_SourceForOverType As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator66 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpTblsPlh_DeleteTable_fast As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTblsPlh_DeleteTable As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_special_AATableFormatting As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents tbPlh_mnu_convertPlhToHalfPage As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpTbl_mnu_AAPlh_To_HalfPlh_Left As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTbl_mnu_AAPlh_To_HalfPlh_Right As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator67 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpTbl_mnu_AAPlh_Reset_to_FullColumn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tbPlh_mnu_rapidFormat As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpTblsPlh_rapidFormat As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTblsPlh_rapidFormat_Encapsulated As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator68 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpBoxes_mnu_rapidFormat_StdTbl_Force As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpBoxes_mnu_rapidFormat_EncapTbl_Force As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpBoxes_mnu_rapidFormat_StdTbl_Force_LT As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_mnu_rapidFormat_StdTbl_Force_ES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_mnu_rapidFormat_StdTbl_Force_Body As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_mnu_rapidFormat_StdTbl_Force_AP As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_mnu_rapidFormat_EncapTbl_Force_LT As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_mnu_rapidFormat_EncapTbl_Force_ES As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_mnu_rapidFormat_EncapTbl_Force_Body As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBoxes_mnu_rapidFormat_EncapTbl_Force_AP As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpAATbls_mnu_editColumns As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpAATbls_mnu_editRows As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpAATbls_mnu_AATableactions As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Separator69 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Separator70 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Separator71 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grp_Plh_TableColumns_mnu_more As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpTblsEdit_InsertColumnRight As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTblsEdit_InsertColumnLeft As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTblsEdit_Delete_Column As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator72 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpTblsEdit_InsertRowAbove As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTblsEdit_InsertRowBelow As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator73 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpTblsEdit_Delete_Row As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTblsEdit_CopyTable As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTblsEdit_PastePriorTable As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTblsEdit_UndoTableAction As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator74 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Separator75 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpTblsEdit_Convert_EncapsToStd As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTblsEdit_Convert_StdToEncaps As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator76 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpTblsEdit_Split_Table As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_floatingPlaceholders As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpReport_PlH_Handling As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpReport_PlH_convertToInline_findAllFloatingTables_2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_PlH_convertToInline As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_PlH_LockToTop As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_PlHFloat_lock_toMarginsLeftAndTop As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_PlH_LockToParagraph As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_PlH_LockToParagraphAndColumn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator77 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpReport_PlH_FloatEdgeToEdge As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_PlH_FloatWide As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_PlH_FloatMarginToMargin As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_PlH_ColumnWidth As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_PlH_TwoColumnWidth As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_Plh_miscPlaceholders As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpPicts_PasteAsPic As Microsoft.Office.Tools.Ribbon.RibbonButton
    Public WithEvents grpEquations_Numbered As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_mnu03 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grp_waterMark_mnu01 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grp_waterMark_mnu02 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents tabFin_mnu_PageNumFormatting As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents tabFin_mnu_PgNumMgmnt_ReNum As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grp_Finalise_mnu01 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents tabFin_mnu_AccessibilityTools As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpReport_PlH_convertToInline_findAllFloatingTables As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_cabinet_add As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_commercial_add As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_confidential_add As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_restricted_add As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_atg_UNOFFICIAL_add As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_atg_OFFICIAL_add As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_atg_OFFICIAL_Sensitive_add As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_removeAll As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_mnu04 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grp_waterMark_removeSec As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_removeStat As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_mnu05 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grp_waterMark_removeSec_fromSect As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_removeStat_fromSect As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator78 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grp_waterMark_draft_add As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_draftOnly_add As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator79 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grp_waterMark_colour_red_stat As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_colour_grey_stat As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator80 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grp_waterMark_forceStat_StyleToDefault As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_submnu01 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grp_waterMark_bold_sec As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_NOTbold_sec As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_colour_red_sec As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_colour_grey_sec As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_alignment_Centre_sec As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_alignment_Right_sec As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_waterMark_forceSec_StyleToDefault As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator84 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Separator83 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Separator82 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Separator81 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpFixes_ApplyEsNumbering As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpFixes_ApplyStdNumbering As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpFixes_ApplyAppNumbering As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator86 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpFixes_ContinueNumbering As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpFixes_RestartNumbering As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator85 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpFixes_getNumberingDialog As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_PgNumMgmnt_ReNum_std As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_PgNumMgmnt_ReNum_2Part As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_Finalise_CrossRefError As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_Finalise_DoAll As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_Finalise_AllFunctions As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grp_Finalise_upDateCopyrightNotice As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator87 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grp_Finalise_updateFields As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_Finalise_setFootersToBold As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_Finalise_RefreshTOC As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_Finalise_CrossRefFlds_setToArialNarrow As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_Finalise_CrossRefFlds_setRefFldNotBold As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpFixes_Repairs_delSpace1_betweenWords As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpFixes_Repairs_delSpace1_atSentenceEnd As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpWCAG_convertThisDoc As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpWCAG_mnu_ContrastControl As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpWCAG_tool_tableHeaderColour_all As Microsoft.Office.Tools.Ribbon.RibbonButton
    Public WithEvents grpWCAG_notesOnAccessibility As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator88 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Separator89 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents grpWCAG_mnu_SetTransparency As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpWCAG_mnu_SetTransparency_to_0 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpWCAG_mnu_SetTransparency_to_25 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpWCAG_mnu_SetTransparency_to_50 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpWCAG_mnu_SetTransparency_to_75 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpWCAG_mnu_SetTransparency_to_100 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpWCAG_tool_convertAllStyles_toBlack As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_Fixes As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpFixes_Repairs As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpFixes_Repairs_remCharChar As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpFixes_Repairs_remSpaces_indrCells As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpFixes_Repairs_SetLanguage As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents mnu_Pagination As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpFixes_RePaginate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpFixes_PaginateOff As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpFixes_mnu_Other As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpFixes_ScreenUpdatingOff As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpFixes_ScreenUpdatingOn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents mnu_Fixes_ScreenUpdating As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpTst_LoadFromResources_getStylesGuide As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents mnu_makeStyles As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents btn_styles_makeTableText As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_buildDocuments As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpReport_tbHome_btn_buildPortraitReport As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_tbHome_btn_buildLandscapeReport As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpReport_tbHome_btn_buildAABrief As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tbHome_grpLetter_standaloneMemo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tbHome_grpLetter_standaloneLetter As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator90 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents tbHome_mnu_contactDetails As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents tbHome_mnu_contactDetails_letter As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents tbHome_mnu_contactDetails_memo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tbHome_mnu_contactDetails_letter_Melbourne As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tbHome_mnu_contactDetails_letter_Sydney As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tbHome_mnu_contactDetails_letter_Brisbane As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tbHome_mnu_contactDetails_letter_Canberra As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tbHome_mnu_contactDetails_letter_Perth As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tbHome_mnu_contactDetails_letter_Adelaide As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator91 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents tbHome_btn_ToggleView As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tbHome_mnu_CreateReport As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Separator92 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Separator93 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Separator94 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents tbHome_btn_Help As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btn_update_Fields As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator95 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property rbn_Styles() As rbn_aa_Addin00
        Get
            Return Me.GetRibbon(Of rbn_aa_Addin00)()
        End Get
    End Property
End Class
