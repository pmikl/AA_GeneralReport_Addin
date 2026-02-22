Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports System.Deployment.Application
Imports System.Drawing
'
'rev 01.00  20250830
'
Public Class cGlobals
    Inherits cControlsMgr
    'Public objTools As New cTools()
    Public _glb_strBanner_Std As String
    Public _glb_strBanner_Sht As String
    Public _glb_strBanner_Lnd As String
    '
    '*** switches between 2021 and 2024 templates
    Public _glb_doBanners_is_On = False                             'No longer works (too many changes, always set to false). Switches off/on the standard 'Chapter' banner for the Portrait reports
    Public _glb_doApp_as_HeadingAP = False                           'Allows me to switch between heading styles (Heading (AP) and Heading 6,7,8,9) for the Appendices
    Public _glb_footer_PageNumColWidth As Single = 37.2             'The footer page number column width, which is also the offset... Note that it is also hard coded in various locations
    '                                                               True means that the numbered list Heading (AP) is selected. False means Heading 6,7,8,9
    '                                                               Heading 6 = Heaidng 1 (AP)
    'Colours
    Public _glb_colour_purple_Dark As Long
    Public _glb_colour_purple_Mid As Long
    Public _glb_colour_purple_Light As Long
    Public _glb_colour_UnitsGrey As Long
    Public _glb_colour_TableBorders As Long
    Public _glb_colour_PageNum_Grey As Long
    '
    Public _glb_colour_Recommendation_Purple As Long
    Public _glb_colour_Finding_Purple As Long
    Public _glb_colour_CaseStudy_Grey As Long
    '
    Public _glb_colour_WaterMark_Grey_sec As Long
    Public _glb_colour_WaterMark_Grey_stat As Long
    '
    Public _glb_colour_FigureFill As Drawing.Color
    '
    'Cover Page types
    Public _glb_cpType_TGFilledPattern As String
    Public _glb_cpType_TGEmptyPattern As String
    Public _glb_cpType_TGPicturePattern As String
    Public _glb_cpType_TGFilledPattern_frontColour As String
    '
    Public _glb_header_leftEdge As Single                       'Distance (mm) between left edge of page and left edge of the Header Table
    Public _glb_header_rightEdge As Single                      'Distance (mm) between right edge of page and right edge of the Header Table
    Public _glb_footer_leftEdge As Single                       'Distance (mm) between left edge of page and left edge of the Footer Table
    Public _glb_footer_rightEdge As Single                      'Distance (mm) between right edge of page and right edge of the Header Table
    '
    Public _glb_footer_table_height As Single                   'Nominal footer table height in points
    '
    Private var_glb_tbl_OutDent As Single                       'Table outdent in mm
    Public var_glb_tbl_bottomSpacerRowHeight As Single          'Bottom space Row in points
    '
    Public var_glb_style_tblCaption_Line2_Indent As Single
    '
    Public glb_var_style_rptBodyText As String = "Body Text"
    Public glb_var_style_tblCaptionStyle As String = "Caption"
    Public glb_var_style_tblHeaderStyle As String = "Table column headings"
    Public glb_var_style_tblUnitsStyle As String = "Table units row"
    Public glb_var_style_tblTextBoldStyle As String = "Table side heading 1"
    Public glb_var_style_tblTextStyle As String = "Table text"
    '
    Public glb_var_style_tblTextBoldStyle_small As String = "Table side heading 1 (small)"
    Public glb_var_style_tblTextStyle_small As String = "Table text (small)"
    '
    Public glb_var_style_tblNoteStyle As String = "Note"
    Public glb_var_style_tblSourceStyle As String = "Source"
    Public glb_var_style_tblSpacerStyle As String = "spacer_tbl"
    '
    'Style used in the header row shapes for security status text
    Public glb_var_style_waterMark_sec As String = "aa_waterMarkText_sec"
    Public glb_var_style_waterMark_stat As String = "aa_waterMarkText_stat"
    '
    '
    Public glb_var_TemplatesDir_default As String
    Public glb_var_TemplatesDir_alt As String
    Public glb_var_TemplateFileName As String
    '
    Public glb_var_strWebSiteId As String = "mikl.net.au"         'Allowed values mikl.net.au, acilallen.com.au
    Public glb_var_strSoftwareType As String = "addin"              'Allowed values addin, dotx
    '
    Public Sub New()
        MyBase.New()
        '
        'Used by cHeaderFooterManager.hf_headers_insert (about line 407).. These are the edge
        'settings for the Header/Footer Tagbles of the Primary Page
        '
        Me._glb_header_leftEdge = 12.2          'mm
        Me._glb_header_rightEdge = -1.0         'flush with right margin
        Me._glb_footer_leftEdge = -1.0        'flush with left margin
        Me._glb_footer_rightEdge = -1.0         'flush with right margin
        'Me._glb_footer_rightEdge = 6            ' 6mm form the left edge of the page
        Me._glb_footer_rightEdge = 5            ' 6mm form the left edge of the page
        '
        Me._glb_footer_table_height = 31                'Nominal height of the footer table in points (body of the report)
        '
        'Me.var_glb_tbl_OutDent = 8.0                    'Table outdent in mm... This is taken into account in the section toggleWidth, so table
        Me.var_glb_tbl_OutDent = 0.0                     'Table outdent in mm... outdents don't go past the header table

        Me.var_glb_tbl_bottomSpacerRowHeight = 8.0      'Bottom spacer row in points
        '
        Me._glb_strBanner_Std = "bnr_Std"
        Me._glb_strBanner_Sht = "bnr_Sht"
        Me._glb_strBanner_Lnd = "bnr_Lnd"
        '
        '
        Me._glb_cpType_TGFilledPattern = "cp_TG_filledPattern"
        Me._glb_cpType_TGEmptyPattern = "cp_TG_emptyPattern"
        Me._glb_cpType_TGPicturePattern = "cp_TG_picturePattern"
        Me._glb_cpType_TGFilledPattern_frontColour = "cp_TG_filledPattern_frontColour"
        '
        'Colours
        Me._glb_colour_purple_Dark = RGB(20, 0, 52)          'Back fill Cover Page etc
        Me._glb_colour_purple_Mid = RGB(108, 63, 153)        'Logo underline and View Colour for Light purple text
        Me._glb_colour_purple_Light = RGB(157, 133, 190)     'Text on dark purple background
        Me._glb_colour_UnitsGrey = RGB(229, 229, 229)
        Me._glb_colour_TableBorders = RGB(100, 100, 100)
        Me._glb_colour_FigureFill = Color.FromArgb(233, 233, 233)
        Me._glb_colour_Recommendation_Purple = RGB(216, 206, 229)
        Me._glb_colour_Finding_Purple = RGB(235, 231, 242)
        '
        Me._glb_colour_PageNum_Grey = RGB(149, 149, 149)
        '
        Me._glb_colour_CaseStudy_Grey = RGB(200, 200, 200)
        Me._glb_colour_WaterMark_Grey_sec = RGB(147, 147, 147)
        Me._glb_colour_WaterMark_Grey_stat = RGB(180, 180, 180)

        '
        Me.var_glb_style_tblCaption_Line2_Indent = 65.4                                 'This is the indent of the second line of the caption style
        '
        'Me.var_glb_style_tblCaption_Line2_Indent = 80.4                                'This is the indent of the second line of the caption style
        'Me.glb_var_TemplatesDir_default = "C:\Templates"
        'Me.glb_var_TemplatesDir_default = IO.Path.GetTempPath().TrimEnd("\"c, "/"c) + "aa_Documents"
        Me.glb_var_TemplatesDir_default = IO.Path.GetTempPath() + "aa_Documents"        'IO.Path.GetTempPath() will return something like C:\Users\peter\AppData\Local\Temp\

        Me.glb_var_TemplatesDir_alt = Me.glb_getDir_documentsLocal()
        Me.glb_var_TemplateFileName = "AA GeneralReport.dotx"
        '
        'Me.glb_var_strWebSiteId = "acilallen.com.au"                    'mikl.net.au, acilallen.com.au
        Me.glb_var_strWebSiteId = "mikl.net.au"                        'mikl.net.au, acilallen.com.au
        Me.glb_var_strSoftwareType = "addin"                            'addin, template

    End Sub
    '
    ''' <summary>
    ''' This method allows you to get the software type. This is either 'addin' or 'dotx'
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_get_SoftwareType() As String
        Return glb_var_strSoftwareType
    End Function
    '
    ''' <summary>
    ''' This method allows you to get the web site id. Allowed values are 'acilallen.com.au' or 'mikl.net.au'
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_get_webSiteId() As String
        Return glb_var_strWebSiteId
    End Function
    '
    ''' <summary>
    ''' This method will return the actual directory being used for the templates file
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_getDir_inUseforTemplates() As String
        Dim strDirActual As String
        '
        'strDirActual = Globals.ThisAddIn.strActualDirTemplates
        strDirActual = Globals.ThisAddIn.strActualDirTemplates
        '
        Return strDirActual
        '
    End Function
    '
    ''' <summary>
    ''' This method will return the local (and not the onedrive) version of the Documents directory
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_getDir_documentsLocal() As String
        Dim strLocalDocumentDir As String
        '
        strLocalDocumentDir = System.IO.Path.Combine(Environment.GetEnvironmentVariable("USERPROFILE"), "Documents")
        '
        Return strLocalDocumentDir
        '
    End Function
    '
    ''' <summary>
    ''' This method will check for the existence of the default Templates directory (C:\Templates). If it doesn't exist it will try
    ''' to create it. If the first use of the method 'objFileMgr.file_make_dir' returns true then the directory exists or it's creation was successful.
    ''' Then the method will  ensure that Word's 'WorkGroup templates file location' is set to the default directory... If it does not exists even after
    ''' the attempt to create it then the method will attempt to create the directory '\aa_Templates' in the local (not onedrive) version
    ''' of the 'Documents' folder and will set the WorkGroup templates accordingly.. If none of the creations are successful, then the
    ''' 'Workgroup templates file location' is set to the local (not onedrive) Documents folder
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_setDir_Templates() As String
        Dim objFileMgr As New cFileHandler()
        Dim strTemplatesDir_alt, strDirCurrent As String
        Dim tst As Boolean
        '
        tst = False
        '
        strTemplatesDir_alt = Me.glb_var_TemplatesDir_alt + "\aa_Templates"
        '
        'If objFileMgr.file_make_dir(Me.glb_var_TemplatesDir_default) Then

        If objFileMgr.file_make_dir(Me.glb_var_TemplatesDir_default) Then
            'Globals.ThisDocument.Application.Options.DefaultFilePath(Microsoft.Office.Interop.Word.WdDefaultFilePath.wdWorkgroupTemplatesPath) = Me.glb_var_TemplatesDir_default
            Globals.ThisAddIn.Application.Options.DefaultFilePath(Microsoft.Office.Interop.Word.WdDefaultFilePath.wdWorkgroupTemplatesPath) = Me.glb_var_TemplatesDir_default
            strDirCurrent = Me.glb_var_TemplatesDir_default
        Else
            If objFileMgr.file_make_dir(strTemplatesDir_alt) Then
                'Globals.ThisDocument.Application.Options.DefaultFilePath(Microsoft.Office.Interop.Word.WdDefaultFilePath.wdWorkgroupTemplatesPath) = strTemplatesDir_alt
                Globals.ThisAddIn.Application.Options.DefaultFilePath(Microsoft.Office.Interop.Word.WdDefaultFilePath.wdWorkgroupTemplatesPath) = strTemplatesDir_alt
                strDirCurrent = strTemplatesDir_alt
            Else
                'Globals.ThisDocument.Application.Options.DefaultFilePath(Microsoft.Office.Interop.Word.WdDefaultFilePath.wdWorkgroupTemplatesPath) = Me.glb_var_TemplatesDir_alt
                Globals.ThisAddIn.Application.Options.DefaultFilePath(Microsoft.Office.Interop.Word.WdDefaultFilePath.wdWorkgroupTemplatesPath) = Me.glb_var_TemplatesDir_alt
                strDirCurrent = strTemplatesDir_alt
            End If
        End If
        '
        Return strDirCurrent
    End Function
    '
    ''' <summary>
    ''' This method will return the full path name of the current template
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_getTmpl_FullName() As String
        Dim strFullName As String
        '
        strFullName = Me.glb_getDir_inUseforTemplates() + "\" + Me.glb_var_TemplateFileName
        '
        Return strFullName
    End Function
    '
    ''' <summary>
    ''' This method will set the sepcified style myStyle to multiple line spacing with
    ''' a value of linespacing (e.g. 0.8, 1.2 etc)
    ''' </summary>
    ''' <param name="myStyle"></param>
    ''' <param name="linespacing"></param>
    Public Sub glb_style_setLineSpacing_AsMultiple(ByRef myStyle As Word.Style, linespacing As Decimal)
        '
        With myStyle.ParagraphFormat
            .LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple
            .LineSpacing = glb_get_wrdApp.LinesToPoints(0.8F)
        End With
        '
    End Sub

    ''' <summary>
    ''' This method will get the specified style. If it doesn't exist that style will be created (based
    ''' on Normal) and returned
    ''' </summary>
    ''' <param name="strStyleName"></param>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function glb_styles_getCreate(ByRef strStyleName As String, ByRef myDoc As Word.Document) As Word.Style
        Dim styl, targetStyle As Word.Style
        '
        styl = Nothing
        targetStyle = Nothing
        '
        Try
            targetStyle = myDoc.Styles.Item(strStyleName)
        Catch ex As Exception
            targetStyle = myDoc.Styles.Add(strStyleName, WdStyleType.wdStyleTypeParagraph)
            targetStyle.BaseStyle = "Normal"
        End Try
        '
        Return targetStyle
    End Function
    '
    '
    ''' <summary>
    ''' This method will update all StyleRef fields in the current document's Footers
    ''' </summary>
    Public Sub glb_flds_updateStyleRefsFooters()
        Dim fld As Word.Field
        Dim myDoc As Word.Document
        Dim hf As Word.HeaderFooter
        Dim rng As Word.Range
        '
        Try
            myDoc = Me.glb_get_wrdActiveDoc()
            rng = myDoc.Application.Selection.Range         'Preserve the current selection
            '
            For Each sect In myDoc.Sections
                '
                Try
                    For Each hf In sect.Footers
                        If hf.Exists Then
                            For Each fld In hf.Range.Fields
                                If fld.Type = WdFieldType.wdFieldStyleRef Then fld.Update()
                            Next
                        End If
                    Next
                Catch ex2 As Exception

                End Try
            Next
            '
            rng.Select()                                    'Reselect
        Catch ex As Exception

        End Try

    End Sub

    '
    Public Sub glb_styles_initStyleNames()
        '
        'Public Const glb_style_tblCaptionStyle As String = "Caption"
        'Public Const glb_style_tblHeaderStyle As String = "Table column headings"
        'Public Const glb_style_tblUnitsStyle As String = "Table units row"
        'Public Const glb_style_tblTextBoldStyle As String = "Table side heading 1"
        'Public Const glb_style_tblTextStyle As String = "Table text"
        '
        'Public Const glb_style_tblTextBoldStyle_small As String = "Table side heading 1 (small)"
        'Public Const glb_style_tblTextStyle_small As String = "Table text (small)"
        '
        'Public Const glb_style_tblNoteStyle As String = ""
        'Public Const glb_style_tblSourceStyle As String = ""
        'Public Const glb_style_tblSpacerStyle As String = ""

    End Sub
    '
    ''' <summary>
    ''' This method will trun screen updating on, then it will refresh the screen and then
    ''' turn screen updating off
    ''' </summary>
    Public Sub glb_screen_update(Optional leaveOn As Boolean = False)
        '
        Me.glb_get_wrdApp.ScreenUpdating = True
        Me.glb_get_wrdApp.ScreenRefresh()
        Me.glb_get_wrdApp.ScreenUpdating = leaveOn

    End Sub
    '
    ''' <summary>
    ''' This method will update/ScreenRefresh no matter the state of ScreenUpdating.. It will leave
    ''' the state of ScreenUpdating as it was before entry
    ''' </summary>
    Public Sub glb_screen_updateLeaveAsItWas()
        Dim currentUpdating As Boolean
        '
        currentUpdating = Me.glb_get_wrdApp.ScreenUpdating
        Me.glb_get_wrdApp.ScreenUpdating = True
        Me.glb_get_wrdApp.ScreenRefresh()
        Me.glb_get_wrdApp.ScreenUpdating = currentUpdating
        '
    End Sub
    '
    ''' <summary>
    ''' This method will stop screen refresh
    ''' </summary>
    Public Sub glb_screen_stopRefresh()
        '
        Me.glb_get_wrdApp.ScreenUpdating = True
        Me.glb_get_wrdApp.ScreenRefresh()
        Me.glb_get_wrdApp.ScreenUpdating = False
        '
    End Sub
    '
    ''' <summary>
    ''' This method will start screen refresh
    ''' </summary>
    Public Sub glb_screen_startRefresh()
        '
        Me.glb_get_wrdApp.ScreenUpdating = True
        Me.glb_get_wrdApp.ScreenRefresh()
        '
    End Sub
    '
    ''' <summary>
    ''' This method will start screen refresh
    ''' </summary>
    Public Sub glb_view_setToPrintLayout()
        '
        Try
            Me.glb_get_wrdApp.ActiveWindow.View.Type = WdViewType.wdPrintView
        Catch ex As Exception

        End Try
        '
    End Sub
    '
    '

#Region "Version and Publish Information"

    Public Function glb_get_PublishVersion() As String
        Dim lst As New Collection()
        Dim strPublishVersion As String
        '
        Try
            lst = Me.glb_get_VersionInformation()
            strPublishVersion = CStr(lst("major")) + "." + CStr(lst("minor")) + "."
            strPublishVersion = strPublishVersion + CStr(lst("build")) + "." + CStr(lst("revision"))
            '
            glb_get_PublishVersion = strPublishVersion
        Catch ex As Exception
            glb_get_PublishVersion = "Debug"
        End Try
        '
    End Function
    '
    '
    Public Function glb_get_UpdateSite() As String
        Dim lst As New Collection()
        '
        Try
            lst = Me.glb_get_VersionInformation()
            glb_get_UpdateSite = CStr(lst("updateSite"))
        Catch ex As Exception
            glb_get_UpdateSite = ""
        End Try
        '
    End Function
    '    
    Public Function glb_get_VersionInformation() As Collection
        Dim updateSite As Uri
        Dim strMajor, strMinor, strBuild, strRevision As String
        Dim strRslt As String
        Dim k As Integer
        k = 1
        '
        Dim lst As New Collection()
        '
        strRslt = ""

        If ApplicationDeployment.IsNetworkDeployed Then
            Dim deploy = ApplicationDeployment.CurrentDeployment
            strMajor = CStr(deploy.CurrentVersion.Major)
            strMinor = CStr(deploy.CurrentVersion.Minor)
            strBuild = CStr(deploy.CurrentVersion.Build)
            strRevision = CStr(deploy.CurrentVersion.Revision)
            updateSite = deploy.UpdateLocation
            lst.Add(strMajor, "major")
            lst.Add(strMinor, "minor")
            lst.Add(strBuild, "build")
            lst.Add(strRevision, "revision")
            lst.Add(updateSite.AbsoluteUri, "updateSite")
        Else
            ' Fallback: use assembly version or mark as non-deployed
            Dim asm = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            lst.Add(CStr(asm.Major), "major")
            lst.Add(CStr(asm.Minor), "minor")
            lst.Add(CStr(asm.Build), "build")
            lst.Add(CStr(asm.Revision), "revision")
            lst.Add("Not ClickOnce deployed", "updateSite")
        End If
        '
        GoTo loop1

        '
        strMajor = CStr(ApplicationDeployment.CurrentDeployment.CurrentVersion.Major)
        strMinor = CStr(ApplicationDeployment.CurrentDeployment.CurrentVersion.Minor)
        strBuild = CStr(ApplicationDeployment.CurrentDeployment.CurrentVersion.Build)
        strRevision = CStr(ApplicationDeployment.CurrentDeployment.CurrentVersion.Revision)
        '

        Try
            lst.Add(strMajor, "major")
            lst.Add(strMinor, "minor")
            lst.Add(strBuild, "build")
            lst.Add(strRevision, "revision")
            '
            updateSite = ApplicationDeployment.CurrentDeployment.UpdateLocation
            lst.Add(updateSite.AbsoluteUri, "updateSite")
            '
            'strURISegments = updateSite.Segments()
            'For j = 0 To strURISegments.Count - 2
            'strRslt = strRslt + strURISegments(j)
            'Next
            'lst.Add(strRslt, "updateSite")

        Catch ex As Exception
            lst = New Collection()
        End Try
        '
loop1:
        Return lst
        '
    End Function
    '
    ''' <summary>
    ''' Gets the update site without the vsto information
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_get_SiteInformation() As Collection
        Dim updateSite As Uri
        Dim strRslt As String
        Dim strURISegments() As String
        Dim j As Integer
        '
        Dim lst As New Collection()
        '
        strRslt = ""
        '
        Try
            '
            updateSite = ApplicationDeployment.CurrentDeployment.UpdateLocation
            'lst.Add(updateSite.AbsoluteUri, "updateSite")
            '
            'strTmp = updateSite.
            'lst.Add(updateSite.GetLeftPart(UriPartial.Path), "updateSite")
            '
            strURISegments = updateSite.Segments()
            For j = 0 To strURISegments.Count - 2
                strRslt = strRslt + strURISegments(j)
            Next
            '
            strRslt = updateSite.GetLeftPart(UriPartial.Authority) + strRslt
            lst.Add(strRslt, "updateSite")

        Catch ex As Exception
            lst = New Collection()
        End Try
        '
        Return lst
        '
    End Function

#End Region

    '
    ''' <summary>
    ''' This method will return (in points) the left and right edges of the
    ''' specified header or footer table as identifed in strEdgeType
    ''' ('header_leftEdge', 'header_rightEdge', 'footer_leftEdge', 'footer_rightEdge')
    ''' </summary>
    ''' <param name="strEdgeType"></param>
    ''' <returns></returns>
    Public Function xglb_hfs_getHFTableEdge(strEdgeType As String) As Single
        Dim rslt As Single
        '
        Select Case strEdgeType
            Case "header_leftEdge"
                rslt = 72.0 * (Me._glb_header_leftEdge / 25.4)
            Case "header_rightEdge"
                rslt = 72.0 * (Me._glb_header_rightEdge / 25.4)
            Case "footer_leftEdge"
                rslt = 72.0 * (Me._glb_footer_leftEdge / 25.4)
            Case "footer_rightEdge"
                rslt = 72.0 * (Me._glb_footer_rightEdge / 25.4)
        End Select
        '
        Return rslt
    End Function
    '
    '
    ''' <summary>
    ''' This method will return the Header Table in WdHeaderFooterIndex.wdHeaderFooterPrimary, via the
    ''' referenced variable tbl_header. It will also return the width of the Header Table in pts
    ''' </summary>
    ''' <param name="tbl_header"></param>
    ''' <returns></returns>
    Public Function glb_hfs_getHeaderTable(ByRef tbl_header As Word.Table) As Single
        Dim sect As Word.Section
        Dim hf As Word.HeaderFooter
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        Dim tblWidth As Single
        '
        sect = Me.glb_get_wrdSect()
        hf = sect.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        '
        tbl_header = Nothing
        tblWidth = 0
        '
        If hf.Range.Tables.Count > 0 Then
            tbl_header = hf.Range.Tables.Item(1)
            dr = tbl_header.Rows.Item(1)
            For Each drCell In dr.Cells
                tblWidth = tblWidth + drCell.Width
            Next
        Else
            tbl_header = Nothing
            tblWidth = 0
        End If

        Return tblWidth
    End Function
    '
    '
    ''' <summary>
    ''' This method will return the width of the Header Table in pts
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function glb_hfs_getHeaderTableWidth(ByRef sect As Word.Section, Optional strHeaderType As String = "primary") As Single
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        Dim tbl_header As Word.Table
        Dim tblWidth As Single
        '
        tbl_header = objHFMgr.hf_get_HeaderTable(sect, strHeaderType)
        tblWidth = 0.0
        '
        Try
            dr = tbl_header.Rows.Item(1)
            For Each drCell In dr.Cells
                tblWidth = tblWidth + drCell.Width
            Next
        Catch ex As Exception
            tblWidth = 0.0
        End Try
        '
        Return tblWidth
        '
    End Function
    '
    Public Function glb_hfs_getFooterTable_Height_Nominal() As Single
        Dim height As Single

        height = Me._glb_footer_table_height
        '
        Return height
    End Function
    '
    'Public Function glb_hfs_getFooterTable_Height_Actual() As Single
    'Dim height As Single

    'height = Me._glb_footer_table_height
    '
    'Return height
    'End Function

    '
    ''' <summary>
    ''' This method will return the Header Table in WdHeaderFooterIndex.wdHeaderFooterPrimary, via the
    ''' referenced variable tbl_header. It will also return the width of the Header Table in pts
    ''' </summary>
    ''' <param name="tbl_footer"></param>
    ''' <returns></returns>
    Public Function glb_hfs_getFooterTable(ByRef tbl_footer As Word.Table) As Single
        Dim sect As Word.Section
        Dim hf As Word.HeaderFooter
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        Dim tblWidth As Single
        '
        sect = Me.glb_get_wrdSect()
        hf = sect.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        '
        tbl_footer = Nothing
        tblWidth = 0
        '
        If hf.Range.Tables.Count > 0 Then
            tbl_footer = hf.Range.Tables.Item(1)
            dr = tbl_footer.Rows.Item(1)
            For Each drCell In dr.Cells
                tblWidth = tblWidth + drCell.Width
            Next
        Else
            tbl_footer = Nothing
            tblWidth = 0
        End If

        Return tblWidth
    End Function
    '
    ''' <summary>
    ''' This method will return the Table Width of tbl. It does so in a table type safe way. So
    ''' it will work for regular and non regular tables. Except for those withe a vertically merged
    ''' cell in the first row. That row must have a 'row like' structure, which will be the case
    ''' for aac legacy tables. If the table is a legacy (outdented) AAC table, then the returned width is the
    ''' width of the body of the table (not the first row). The width of the outdented first row will be the 
    ''' table width + 'left padding in the first cell.. If tblWidth could not be deterined
    ''' it returns 0.0
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function glb_tbls_getTableWidth(ByRef tbl As Word.Table) As Single
        Dim drCell As Word.Cell
        Dim tblWidth As Single
        '
        tblWidth = 0.0
        '
        Try
            For Each drCell In tbl.Range.Cells
                If drCell.RowIndex = 1 Then
                    tblWidth = tblWidth + drCell.Width
                End If
            Next
            '
            If Me.glb_tbl_isLegacyAATable(tbl) Then tblWidth = tblWidth - tbl.Range.Cells.Item(1).LeftPadding
        Catch ex As Exception
            tblWidth = 0.0
        End Try
        '
        Return tblWidth
        '
    End Function
    '
    ''' <summary>
    ''' This method will return true if the table (tbl) is an AAC legacy Table. It looks at the padding in the
    ''' first cell, which should be 22.7 pt (It checks that the padding is in a range 15.0 to 223.o pt). Then
    ''' to make sure it also checks that the padding of the second cell (which should normally be 0.0) is less than
    ''' the padding of the first cell... This test is table type safe. It will work for both row regular, column regular
    ''' and Tables with mixed horizontal and vertically merged cells
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function glb_tbl_isLegacyAATable(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim leftPadding, leftPadding2 As Single
        '
        '
        Try
            leftPadding = tbl.Range.Cells.Item(1).LeftPadding
            leftPadding2 = tbl.Range.Cells.Item(2).LeftPadding
            '
            'Standard left Padding in AAC legacy table is 22.7 pt... Also the padding in the second cell 
            'is normall 0.0
            If (leftPadding >= 15 And leftPadding <= 24.0) And leftPadding > leftPadding2 Then
                rslt = True
            End If
            '
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
    End Function
    '
    ''' <summary>
    ''' This method will test the table (tbl) and if it is regular by row (i.e. no
    ''' vertically merged cells it will return true.... 20231126 Works
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function glb_tbls_isRegularByRow(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim dr As Word.Row
        Dim knt As Integer
        Dim str As String
        '
        rslt = False
        '
        Try
            For Each dr In tbl.Rows
                knt = dr.Cells.Count
                str = knt.ToString()
            Next
            rslt = True
        Catch ex As Exception
            rslt = False
        End Try

        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' This method will test the table (tbl) and if it is regular by columns (i.e. no
    ''' horizontally merged cells it will return true.... 20231126 Works
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function glb_tbls_isRegularByCol(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim drCol As Word.Column
        Dim knt As Integer
        Dim str As String
        '
        rslt = False
        '
        Try
            For Each drCol In tbl.Columns
                knt = drCol.Cells.Count
                str = knt.ToString()
            Next
            rslt = True
        Catch ex As Exception
            rslt = False
        End Try

        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' This method will return the last cell of the table (tbl). It does so
    ''' in a Table type safe way
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function glb_tbls_getLastCell(ByRef tbl As Word.Table) As Word.Cell
        Dim drCell As Word.Cell
        '
        drCell = Nothing
        If Not IsNothing(tbl) Then
            drCell = tbl.Range.Cells.Item(tbl.Range.Cells.Count)
        Else
            drCell = Nothing
        End If

        Return drCell

    End Function

    '
    ''' <summary>
    ''' This method returns true if the table (tbl) is both regular by Row and by Column
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function glb_tbls_isRegular(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        '
        rslt = False
        rslt = (Me.glb_tbls_isRegularByCol(tbl)) And (Me.glb_tbls_isRegularByRow(tbl))
        '
        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' This method retrieves the 'aac Table (no lines)' table style. If it doesn't exist it creates it. If doExtraFormatting is true,
    ''' then the existing style is formatted as per the 'doExtraFormatting' code... If the style is created, then the 'doExtraFormatting'
    ''' code is run regardless of the setting of 'doExtraFormatting'
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="doExtraFormatting"></param>
    ''' <returns></returns>
    Public Function glb_tbl_getAACTableNoLinesStyle(ByRef myDoc As Word.Document, Optional doExtraFormatting As Boolean = False) As Word.Style
        Dim tblStyle As Word.Style
        Dim borderColour As Long
        '
        borderColour = RGB(0, 0, 0)
        Try
            tblStyle = myDoc.Styles.Item("aac Table (no lines)")
            tblStyle.BaseStyle = myDoc.Styles.Item("Table Normal")

        Catch ex As Exception
            tblStyle = myDoc.Styles.Add("aac Table (no lines)", WdStyleType.wdStyleTypeTable)
            tblStyle.BaseStyle = myDoc.Styles.Item("Table Normal")
            doExtraFormatting = True
        End Try
        '
        If doExtraFormatting Then
            tblStyle.Table.AllowPageBreaks = True
            tblStyle.Table.AllowBreakAcrossPage = False
            '
            tblStyle.Table.TopPadding = 0.0#
            tblStyle.Table.BottomPadding = 0.0#
            tblStyle.Table.LeftPadding = 0.0#
            tblStyle.Table.RightPadding = 0.0#
            '
            tblStyle.Table.Borders(WdBorderType.wdBorderHorizontal).LineStyle = WdLineStyle.wdLineStyleNone
            '
            tblStyle.Table.Borders(WdBorderType.wdBorderVertical).LineStyle = WdLineStyle.wdLineStyleNone
            tblStyle.Table.Borders(WdBorderType.wdBorderLeft).LineStyle = WdLineStyle.wdLineStyleNone
            tblStyle.Table.Borders(WdBorderType.wdBorderRight).LineStyle = WdLineStyle.wdLineStyleNone
            '
            tblStyle.Table.Borders(WdBorderType.wdBorderTop).LineStyle = WdLineStyle.wdLineStyleNone
            '
            tblStyle.Table.Borders(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleNone
            'tblStyle.Table.Borders(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth050pt
            'tblStyle.Table.Borders(WdBorderType.wdBorderBottom).Color = borderColour
            '
            tblStyle.Table.Alignment = WdRowAlignment.wdAlignRowLeft
            '
            '
        End If
        '
        Return tblStyle

    End Function
    '
    ''' <summary>
    ''' This method will add the table style 'aac Table (Basic)' to myDoc if it doesn't exist. If
    ''' it does it applies the appropriate formatting, which gives us an 'in' to changing this
    ''' construct at a later date. The doExtraFormatting option is applied to an existing style
    ''' if doExtraFormatting is true. It is always applied (regardless of the option setting)
    ''' to the style if it is newly created.... 
    ''' 
    ''' Verified 20231201 (just a problem with the line under the first row.. At the moment taken off
    ''' post production
    ''' 
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="doExtraFormatting"></param>
    Public Function glb_tbl_getAACTableBasicStyle(ByRef myDoc As Word.Document, Optional doExtraFormatting As Boolean = False) As Word.Style
        Dim tblStyle As Word.Style
        Dim borderColour As Long
        Dim strText As String
        '
        'borderColour = Me._glb_colour_purple_Dark
        borderColour = Me._glb_colour_TableBorders
        'borderColour = RGB(255,0,0)

        '
        Try
            tblStyle = myDoc.Styles.Item("aac Table (Basic)")
            strText = tblStyle.NameLocal
            '
        Catch ex As Exception
            tblStyle = myDoc.Styles.Add("aac Table (Basic)", WdStyleType.wdStyleTypeTable)
            tblStyle.BaseStyle = myDoc.Styles.Item("Table Normal")
            doExtraFormatting = True
        End Try
        '
        'tbl.AllowPageBreaks = True                                           'Will allow a row to break across pages
        'tbl.Rows.AllowBreakAcrossPages = False
        '
        If doExtraFormatting Then
            tblStyle.Table.AllowPageBreaks = True
            tblStyle.Table.AllowBreakAcrossPage = False
            tblStyle.Table.Alignment = WdRowAlignment.wdAlignRowLeft
            'tblStyle.Table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed)              'columns sizes don't change to accommodate text, will  AllowAutoFit = false

            '
            tblStyle.Table.TopPadding = 0.0#
            tblStyle.Table.BottomPadding = 0.0#
            tblStyle.Table.LeftPadding = 0.0#
            tblStyle.Table.RightPadding = 0.0#
            '
            tblStyle.Table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleNone
            tblStyle.Table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone
            '
            tblStyle.Table.Borders(WdBorderType.wdBorderHorizontal).LineStyle = WdLineStyle.wdLineStyleSingle
            tblStyle.Table.Borders(WdBorderType.wdBorderHorizontal).LineWidth = WdLineWidth.wdLineWidth050pt
            tblStyle.Table.Borders(WdBorderType.wdBorderHorizontal).Color = borderColour
            '
            tblStyle.Table.Borders(WdBorderType.wdBorderVertical).LineStyle = WdLineStyle.wdLineStyleNone
            tblStyle.Table.Borders(WdBorderType.wdBorderLeft).LineStyle = WdLineStyle.wdLineStyleNone
            tblStyle.Table.Borders(WdBorderType.wdBorderRight).LineStyle = WdLineStyle.wdLineStyleNone
            '
            tblStyle.Table.Borders(WdBorderType.wdBorderTop).LineStyle = WdLineStyle.wdLineStyleNone
            'tblStyle.Table.Borders(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleNone
            '
            tblStyle.Table.Borders(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleSingle
            tblStyle.Table.Borders(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth050pt
            tblStyle.Table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone
            '
            tblStyle.Table.Condition(WdConditionCode.wdFirstRow).Shading.ForegroundPatternColor = Me._glb_colour_purple_Dark
            tblStyle.Table.Condition(WdConditionCode.wdFirstRow).Borders(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleNone
            tblStyle.Table.Condition(WdConditionCode.wdFirstRow).Borders(WdBorderType.wdBorderBottom).Visible = False
            '

        End If
        '
        '
        Return tblStyle
        '
    End Function
    '

    Public Sub glb_tbl_apply_aacTableNoLinesStyle(ByRef tbl As Word.Table)
        Dim myDoc As Word.Document
        Dim objStylesMgr As New cStylesManager()
        Dim objTblStyles As New cTableStyles()
        'Dim tblStyle As Word.Style
        ' tblstyl_add_aacTableNoLines

        myDoc = tbl.Range.Document
        tbl.Style = Me.glb_tbl_getAACTableNoLinesStyle(myDoc)
        tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed)
        '
        tbl.Range.Style = objStylesMgr.style_txt_getTableTextStyle(myDoc)
        tbl.Rows.First.Range.Style = objStylesMgr.style_txt_getTableHeadingStyle(myDoc)
        '
        tbl.AllowPageBreaks = True
        tbl.Rows.AllowBreakAcrossPages = False
        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
        tbl.PreferredWidth = 100
        '
    End Sub
    '

    Public Sub glb_tbl_apply_aacTableBasicStyle(ByRef tbl As Word.Table)
        Dim myDoc As Word.Document
        Dim objStylesMgr As New cStylesManager()
        Dim objTblStyles As New cTableStyles()
        Dim headingStyle As Word.Style
        'Dim tblStyle As Word.Style
        ' tblstyl_add_aacTableNoLines

        myDoc = tbl.Range.Document
        tbl.Style = Me.glb_tbl_getAACTableBasicStyle(myDoc)
        tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed)
        '
        tbl.Range.Style = objStylesMgr.style_txt_getTableTextStyle(myDoc)
        '
        tbl.Rows.LeftIndent = 0
        '
        tbl.AllowPageBreaks = True
        tbl.Rows.AllowBreakAcrossPages = False
        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
        tbl.PreferredWidth = 100
        '
        '
        tbl.ApplyStyleHeadingRows = True
        headingStyle = objStylesMgr.style_txt_getTableHeadingStyle(myDoc)
        tbl.Rows.First.Range.Style = headingStyle
        tbl.Rows.First.Range.Font.Color = headingStyle.Font.Color
        '
        'tbl.Rows.First.HeadingFormat = True
        '
        '
    End Sub


    ''' <summary>
    ''' This method will autofit a standard table (no horizontally merged cells or top row offset) between the margins.
    ''' It does so by expanding column widths
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="deleteBackground"></param>
    Public Sub glb_tbls_AutoFitBanner(ByRef tbl As Word.Table, Optional deleteBackground As Boolean = False)
        Dim tblWidth, widthBetweenMargins As Single
        Dim shp As Word.Shape
        Dim drCell As Word.Cell
        Dim drCol As Word.Column
        Dim sect As Word.Section
        Dim delta As Single
        '
        delta = 0.0
        sect = tbl.Range.Sections.Item(1)
        widthBetweenMargins = glb_get_widthBetweenMargins(sect)
        '
        tblWidth = Me.glb_tbls_getTableWidth(tbl)
        '
        If tblWidth <> 0.0 Then
            delta = (widthBetweenMargins - tblWidth) / tbl.Columns.Count
            For Each drCol In tbl.Columns
                drCol.Width = drCol.Width + delta
            Next
            '
            drCell = tbl.Range.Cells.Item(1)
            Try
                If drCell.Range.ShapeRange.Count <> 0 Then
                    If deleteBackground Then
                        shp = drCell.Range.ShapeRange.Item(1)
                        shp.Delete()
                    Else
                        shp = drCell.Range.ShapeRange.Item(1)
                        shp.Width = tbl.Range.Columns.Item(1).Width + tbl.Range.Columns.Item(2).Width
                    End If
                End If
                '
            Catch ex As Exception

            End Try
        End If
        '
finis:
        '
    End Sub '
    '
    ''' <summary>
    ''' This method will autofit (i.e. between the section margins) the regular table tbl.
    ''' defaults to false
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub glb_tbls_AutoFitRegularTable(ByRef tbl As Word.Table)
        Dim sect As Word.Section
        Dim delta As Single
        '
        delta = 0.0
        sect = tbl.Range.Sections.Item(1)
        'widthBetweenMargins = glb_get_widthBetweenMargins(sect)
        '
        Try
            'tblWidth = Me.glb_tbls_getTableWidth(tbl)
            tbl.PreferredWidth = Me.glb_get_widthBetweenMargins(sect)

            '
            'If tblWidth <> 0.0 Then
            'delta = (widthBetweenMargins - tblWidth) / tbl.Columns.Count
            ' For Each drCol In tbl.Columns
            'drCol.Width = drCol.Width + delta
            'Next
            '
            'End If

        Catch ex As Exception

        End Try
        '
    End Sub

    '
    ''' <summary>
    ''' This method will autofit (i.e. between the section margins) the regular table tbl.
    ''' defaults to false
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub glb_tbls_AutoFitRegularTableToSize(ByRef tbl As Word.Table, tblWidth As Single)
        Dim sect As Word.Section
        Dim delta As Single
        '
        delta = 0.0
        sect = tbl.Range.Sections.Item(1)
        'widthBetweenMargins = glb_get_widthBetweenMargins(sect)
        '
        Try
            'tblWidth = Me.glb_tbls_getTableWidth(tbl)
            tbl.PreferredWidth = tblWidth

            '
            'If tblWidth <> 0.0 Then
            'delta = (widthBetweenMargins - tblWidth) / tbl.Columns.Count
            ' For Each drCol In tbl.Columns
            'drCol.Width = drCol.Width + delta
            'Next
            '
            'End If

        Catch ex As Exception

        End Try
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will autofit (i.e. between the section margins) either all of the regular tables in the
    ''' section sect, or just the first table in sect.. This depends on the value of doAllTables which
    ''' defaults to false
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="doAllTables"></param>
    Public Function glb_tbls_AutoFitRegularTable(ByRef sect As Word.Section, Optional doAllTables As Boolean = False) As Word.Table
        Dim tbl As Word.Table
        Dim delta As Single
        '
        delta = 0.0
        tbl = Nothing
        '
        Try
            If doAllTables Then
                For Each tbl In sect.Range.Tables
                    Me.glb_tbls_AutoFitRegularTable(tbl)
                Next
            Else
                If sect.Range.Tables.Count <> 0 Then
                    tbl = sect.Range.Tables.Item(1)
                    Me.glb_tbls_AutoFitRegularTable(tbl)
                End If
            End If
        Catch ex As Exception
            tbl = Nothing
        End Try
        '
        Return tbl
        '
    End Function

    '
    '
    ''' <summary>
    ''' This method sets the left indent of the table relative to the left margin. So a 
    ''' negative left indent has the table extending to the left of the left margin
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="leftIndent"></param>
    Public Sub glb_tbls_setLeftIndent(ByRef tbl As Word.Table, leftIndent As Single)
        Dim dr As Word.Row
        Dim outDentValue As Single
        '
        outDentValue = Me.glb_tbls_getTableOutDent(tbl)
        '
        If outDentValue = 0.0 Then
            For Each dr In tbl.Rows
                dr.LeftIndent = leftIndent
            Next
        Else
            For Each dr In tbl.Rows
                If dr.Index = 1 Then
                    dr.LeftIndent = leftIndent

                Else
                    dr.LeftIndent = leftIndent - outDentValue
                End If
            Next

        End If
        '
    End Sub
    '
    Public Function glb_tbls_getTableOutDent(ByRef tbl As Word.Table) As Single
        Dim outDentValue As Single
        Dim drHeader, drNext As Word.Row
        '
        drHeader = tbl.Rows.Item(1)
        drNext = tbl.Rows.Item(2)
        outDentValue = drHeader.LeftIndent - drNext.LeftIndent
        '
        Return outDentValue
    End Function
    '    '
    ''' <summary>
    ''' This method will return (in points) the actual left and right edges of the
    ''' specified header or footer table as identifed in strEdgeType
    ''' ('header_leftEdge', 'header_rightEdge', 'footer_leftEdge', 'footer_rightEdge')
    ''' </summary>
    ''' <param name="strEdgeType"></param>
    ''' <returns></returns>
    Public Function glb_hfs_getHFTableEdge(ByRef sect As Word.Section, strEdgeType As String) As Single
        Dim leftEdge, rightEdge, rslt As Single
        Dim objTools As New cTools()
        '
        Select Case strEdgeType
            Case "header_leftEdge"
                leftEdge = Me._glb_header_leftEdge
                If leftEdge < 0.0 Then
                    leftEdge = sect.PageSetup.LeftMargin
                Else
                    leftEdge = objTools.tools_math_MillimetersToPoints(Me._glb_header_leftEdge)
                End If
                rslt = leftEdge
                '
            Case "header_rightEdge"
                rightEdge = Me._glb_header_rightEdge
                If rightEdge < 0.0 Then
                    rightEdge = sect.PageSetup.RightMargin
                Else
                    rightEdge = objTools.tools_math_MillimetersToPoints(Me._glb_header_rightEdge)
                End If
                rslt = rightEdge
                '
            Case "footer_leftEdge"
                leftEdge = Me._glb_footer_leftEdge
                If leftEdge < 0.0 Then
                    leftEdge = sect.PageSetup.LeftMargin
                Else
                    leftEdge = objTools.tools_math_MillimetersToPoints(Me._glb_footer_leftEdge)
                End If
                rslt = leftEdge
                '
            Case "footer_rightEdge"
                rightEdge = Me._glb_footer_rightEdge
                If rightEdge < 0.0 Then
                    rightEdge = sect.PageSetup.RightMargin
                Else
                    rightEdge = objTools.tools_math_MillimetersToPoints(Me._glb_footer_rightEdge)
                End If
                rslt = rightEdge
        End Select
        '
        Return rslt
    End Function
    '
    '
    '
    ''' <summary>
    ''' This function will return the Table outdent in points
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_get_TableOutdent() As Single
        Dim objTools As New cTools()

        Return objTools.tools_math_MillimetersToPoints(Me.var_glb_tbl_OutDent)
    End Function
    '   '
    ''' <summary>
    ''' This function will return the Table outdent in mm
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_get_TableOutdent_mm() As Single
        Return Me.var_glb_tbl_OutDent
    End Function
    '
    '
#Region "Dimensions"
    '
    ''' <summary>
    ''' This function will return the width (in pts) between margins of the specified section
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function glb_get_widthBetweenMargins(ByRef sect As Word.Section) As Single
        Dim width As Single
        '
        width = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
        '
        Return width
        '
    End Function

    '
    'This is here so that the page measurements are in one place. This
    'is the left margin for the cover page, which represents the widest
    'a page can be set. This is used in the widthOfSection_Wide procedure
    Public Function leftMarginAbsoluteEdge() As Single
        leftMarginAbsoluteEdge = 23.8
        leftMarginAbsoluteEdge = 35.45          'T and G version


    End Function
    '
    ''' <summary>
    ''' This method will return the dimensions of the standard Contacts Page in pts
    ''' (topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance)
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_getDimensions_Contacts_Prt(strContactsType As String) As Collection
        Dim lstOfDimensions As New Collection()
        Dim objRptMgr As New cReport()
        '
        'leftEdge = Me.glb_math_MillimetersToPoints(Me._glb_header_leftEdge)
        '
        'topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance
        '
        Select Case strContactsType
            Case "front"
                Call Me.initPageSettings(lstOfDimensions, 56.0#, 36.0, 25.0, 36.0, 0.0#, 22.7, 14.4)              'Long and Short Report Front Contacts Page
            Case "back"
                Call Me.initPageSettings(lstOfDimensions, 56.0#, 36.0, 25.0, 36.0, 0.0#, 22.7, 14.4)              'Long and Short Report Back Contacts Page
        End Select
        '
        'If strContactsType = "front" Then
        'Select Case objRptMgr.Rpt_Mode_Get()
        'Case objRptMgr.modeLong, objRptMgr.modeShort
        'Call Me.initPageSettings(lstOfDimensions, 56.0#, 56.0, 25.0, 36.0, 0.0#, 22.7, 14.4)              'Long and Short Report Front Contacts Page
        'Case objRptMgr.modeLongLandscape
        'Call Me.initPageSettings(lstOfDimensions, 56.0#, 35.0, 25.0, 36.0, 0.0#, 22.7, 14.4)             'Landscape Report Contacts Page
        'Case Else
        'Call Me.initPageSettings(lstOfDimensions, 56.0#, leftEdge, 25.0, 36.0, 0.0#, 22.7, 14.4)              'Default the same as Long and Short Report
        'End Select
        'End If
        '
        '
        'If strContactsType = "back" Then
        'Select Case objRptMgr.Rpt_Mode_Get()
        'Case objRptMgr.modeLong, objRptMgr.modeShort
        'Case objRptMgr.modeLongLandscape
        ' Call Me.initPageSettings(lstOfDimensions, 56.0#, 35.0, 25.0, 36.0, 0.0#, 22.7, 14.4)             'Landscape Report Back Contacts Page
        'Case Else
        'Call Me.initPageSettings(lstOfDimensions, 56.0#, leftEdge, 25.0, 36.0, 0.0#, 22.7, 14.4)              'Default the same as Long and Short Report
        'End Select
        ' End If
        '
        '
        Return lstOfDimensions
        '
    End Function
    '
    ''' <summary>
    ''' This method will show/create the color Picker.. It does so in a way to avoid
    ''' multiple instantiation.. We can access the restricted version frm_colorPicker02
    ''' or the expanded version frm_colorPicker
    ''' </summary>
    ''' <param name="strColourPickerMode"></param>
    Public Sub glb_show_ColorPicker(strColourPickerMode As String)
        '
        Try
            Globals.ThisAddIn.point_PriorClick = System.Windows.Forms.Cursor.Position
            If IsNothing(Globals.ThisAddIn.frm_colorPicker02) Then
                Globals.ThisAddIn.frm_colorPicker02 = New frm_colorPicker02(strColourPickerMode)
                Globals.ThisAddIn.frm_colorPicker02.Show()
                'frmPicker = New frm_colorPicker(strColourPickerMode)
            Else
                Globals.ThisAddIn.frm_colorPicker02.Activate()
                Globals.ThisAddIn.frm_colorPicker02.frm_colorPicker_Rename(strColourPickerMode)
            End If
            '
            Globals.ThisAddIn.frm_colorPicker02.TopMost = True
            Globals.ThisAddIn.frm_colorPicker02.Top = System.Windows.Forms.Cursor.Position.Y + 5
            Globals.ThisAddIn.frm_colorPicker02.Left = System.Windows.Forms.Cursor.Position.X
            '
        Catch ex As Exception
            MsgBox("The custom colour picker could not be built/activated")
        End Try

    End Sub
    '
    ''' <summary>
    ''' This method will return a collection that returns the Application's
    ''' current theme colours as RGB (i.e. 32 bit Integer in VB.NET). In VBA,
    ''' RGB Colours have to be held in type Long. The Integer in VBA is 16 bits only.
    ''' 
    ''' The colours can be access by a key that starts at '0' adn extends (generally)
    ''' to 11.. But this method is not limited to 12 items. Hence the Collection
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_get_docThemeColours_Actual() As Collection
        Dim themeColours As Collection
        Dim thm As OfficeTheme
        Dim colorScheme As ThemeColorScheme
        Dim thm1, thm2, thm3, thm4, thm5, thm6 As ThemeColor
        Dim thm7, thm8, thm9, thm10, thm11, thm12 As ThemeColor
        Dim transparency As Integer
        'Dim objGlobals As New cGlobals()
        '
        transparency = 0
        themeColours = New Collection
        '
        Try
            thm = Me.glb_get_wrdActiveDoc.DocumentTheme
            'thm = Globals.ThisAddIn.Application.ActiveDocument.DocumentTheme
            'thm = Globals.ThisDocument.Application.ActiveDocument.Theme
            colorScheme = thm.ThemeColorScheme
            'For Each thmcolor In colorScheme.Colors
            thm1 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeDark1)
            thm2 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeLight1)
            thm3 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeDark2)
            thm4 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeLight2)
            thm5 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent1)
            thm6 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent2)
            thm7 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent3)
            thm8 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent4)
            thm9 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent5)
            thm10 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent6)
            thm11 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeHyperlink)
            thm12 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeFollowedHyperlink)
            '
            'To cater for the apparent transposiiton when displayed by excel
            '1>0, 0>1, 2>3, 3>2
            themeColours.Add(thm1.RGB, "1")
            themeColours.Add(thm2.RGB, "0")
            themeColours.Add(thm3.RGB, "3")
            themeColours.Add(thm4.RGB, "2")
            themeColours.Add(thm5.RGB, "4")
            themeColours.Add(thm6.RGB, "5")
            themeColours.Add(thm7.RGB, "6")
            themeColours.Add(thm8.RGB, "7")
            themeColours.Add(thm9.RGB, "8")
            themeColours.Add(thm10.RGB, "9")
            themeColours.Add(thm11.RGB, "10")
            themeColours.Add(thm12.RGB, "11")
            '
            'themeColours.Add(thm1.RGB, "0")
            'themeColours.Add(thm2.RGB, "1")
            'themeColours.Add(thm3.RGB, "2")
            'themeColours.Add(thm4.RGB, "3")
            'themeColours.Add(thm5.RGB, "4")
            'themeColours.Add(thm6.RGB, "5")
            'themeColours.Add(thm7.RGB, "6")
            'themeColours.Add(thm8.RGB, "7")
            'themeColours.Add(thm9.RGB, "8")
            'themeColours.Add(thm10.RGB, "9")
            'themeColours.Add(thm11.RGB, "10")
            'themeColours.Add(thm12.RGB, "11")
            '
        Catch ex As Exception

        End Try


        Return themeColours

        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will return the dimensions of a standard report Contacts Landscape Page in pts
    ''' (topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance)
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_getDimensions_Contacts_Lnd() As Collection
        Dim lstOfDimensions As New Collection()
        '
        'Me.glb_hfs_getHFTableEdge("header_leftEdge")
        'leftedge = Me.glb_math_MillimetersToPoints(Me._glb_header_leftEdge)
        '
        'Call Me.initPageSettings(lstOfDimensions, 56.7#, 155.95, 56.7#, 42.55, 0.0#, 22.7, 7.95)
        'Call Me.initPageSettings(lstOfDimensions, 56.7#, leftedge, 56.7#, 42.55, 0.0#, 22.7, 7.95)
        Call Me.initPageSettings(lstOfDimensions, 56.0#, 42.0, 24.0#, 42.0, 0.0#, 22.7, 7.95)

        '
        Return lstOfDimensions
        '
    End Function
    '
    ''' <summary>
    ''' This method will return the dimensions of the standard ReportPortrait Page in pts
    ''' (topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance)
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_getDimensions_Std_Prt() As Collection
        Dim lstOfDimensions As New Collection()
        '
        '(topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance)
        'Call Me.initPageSettings(lstOfDimensions, 56.7#, 155.95, 56.7#, 42.55, 0.0#, 22.7, 7.95)                'Portrait Report 15mm right margin
        '
        '*** Test settings
        'Call Me.initPageSettings(lstOfDimensions, 56.0#, 120, 56.0#, 30, 0.0#, 22.7, 7.95)                'Portrait Report 15mm right margin
        'Call Me.initPageSettings(lstOfDimensions, 56.0#, 120, 76.0#, 50.0, 0.0#, 22.7, 7.95)                'As per sepc 20231204
        'Call Me.initPageSettings(lstOfDimensions, 56.0#, 120, 66.0#, 50.0, 0.0#, 22.7, 7.95)                'As per sepc 20231204
        'Call Me.initPageSettings(lstOfDimensions, 56.0#, 80, 66.0#, 42.0, 0.0#, 22.7, 7.95)                'As per sepc 20231204
        Call Me.initPageSettings(lstOfDimensions, 56.0#, 120, 66.0#, 42.0, 0.0#, 22.7, 7.95)                'As per sepc 20231204

        '
        Return lstOfDimensions
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will return the dimensions of a standard report Landscape Page in pts
    ''' (topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance)
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_getDimensions_Std_Lnd() As Collection
        Dim lstOfDimensions As New Collection()
        '
        'Call Me.initPageSettings(lstOfDimensions, 56.7#, 155.95, 56.7#, 42.55, 0.0#, 22.7, 7.95)
        'Call Me.initPageSettings(lstOfDimensions, 56.0#, 120, 56.0#, 50.0#, 0.0#, 22.7, 7.95)
        Call Me.initPageSettings(lstOfDimensions, 56.0#, 120, 56.0#, 42.0#, 0.0#, 22.7, 7.95)

        '
        Return lstOfDimensions
        '
    End Function
    '

    '
    ''' <summary>
    ''' This method will return the dimensions of the standard ReportPortrait Page in pts
    ''' (topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance)
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_getDimensions_CaseStudy_Prt() As Collection
        Dim lstOfDimensions As New Collection()
        '
        Dim leftedge As Single
        '
        'Me.glb_hfs_getHFTableEdge("header_leftEdge")
        leftedge = Me.glb_math_MillimetersToPoints(Me._glb_header_leftEdge)

        '(topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance)
        'Call Me.initPageSettings(lstOfDimensions, 56.7#, 155.95, 56.7#, 42.55, 0.0#, 22.7, 7.95)                'Portrait Report 15mm right margin
        '
        '*** Test settings
        'Call Me.initPageSettings(lstOfDimensions, 56.0#, 120, 56.0#, 30, 0.0#, 22.7, 7.95)                'Portrait Report 15mm right margin
        'Call Me.initPageSettings(lstOfDimensions, 56.0#, 120, 76.0#, 50.0, 0.0#, 22.7, 7.95)                'As per sepc 20231204
        Call Me.initPageSettings(lstOfDimensions, 56.0#, leftedge, 66.0#, 42.0, 0.0#, 22.7, 7.95)                'As per sepc 20231204

        '
        Return lstOfDimensions
        '
    End Function
    '

    '
    ''' <summary>
    ''' This method will return the dimensions of a standard report Landscape Page in pts
    ''' (topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance)
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_getDimensions_CaseStudy_Lnd() As Collection
        Dim lstOfDimensions As New Collection()
        Dim leftedge As Single
        '
        'Me.glb_hfs_getHFTableEdge("header_leftEdge")
        leftedge = Me.glb_math_MillimetersToPoints(Me._glb_header_leftEdge)
        '
        'Call Me.initPageSettings(lstOfDimensions, 56.7#, 155.95, 56.7#, 42.55, 0.0#, 22.7, 7.95)
        Call Me.initPageSettings(lstOfDimensions, 56.0#, leftedge, 56.0#, 50.0#, 0.0#, 22.7, 7.95)

        '
        Return lstOfDimensions
        '
    End Function


    Public Function glb_getDimensions_Divider_Prt() As Collection
        Dim lstOfDimensions As New Collection()
        '
        '(topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance)
        'Call Me.initPageSettings(lstOfDimensions, 138.0#, 56.0, 56.0#, 36.0, 0.0#, 22.7, 7.95)                'Portrait Report 15mm right margin
        Call Me.initPageSettings(lstOfDimensions, 220.0#, 56.0, 56.0#, 36.0, 0.0#, 22.7, 7.95)                'Portrait Report 15mm right margin
        '
        Return lstOfDimensions
        '
    End Function
    '
    ''' <summary>
    ''' This method will return the dimensions of a standard report Contacts Landscape Page in pts
    ''' (topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance)
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_getDimensions_Divider_Lnd() As Collection
        Dim lstOfDimensions As New Collection()
        Dim leftedge As Single
        '
        leftedge = Me.glb_math_MillimetersToPoints(Me._glb_header_leftEdge)
        '
        'Call Me.initPageSettings(lstOfDimensions, 56.7#, 155.95, 56.7#, 42.55, 0.0#, 22.7, 7.95)
        'Call Me.initPageSettings(lstOfDimensions, 56.7#, leftedge, 56.7#, 42.55, 0.0#, 22.7, 7.95)
        'Call Me.initPageSettings(lstOfDimensions, 56.0#, leftedge, 24.0#, 42.55, 0.0#, 22.7, 7.95)
        'Call Me.initPageSettings(lstOfDimensions, 56.0#, 56.0#, 24.0#, 56.0#, 0.0#, 22.7, 7.95)
        'Call Me.initPageSettings(lstOfDimensions, 138.0#, 56.0#, 24.0#, 36.0#, 0.0#, 22.7, 7.95)
        'Call Me.initPageSettings(lstOfDimensions, 80.0#, 56.0#, 24.0#, 36.0#, 0.0#, 22.7, 7.95)
        Call Me.initPageSettings(lstOfDimensions, 150.0#, 56.0#, 24.0#, 36.0#, 0.0#, 22.7, 7.95)

        '
        Return lstOfDimensions
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will return the dimensions of the standard CoverPage Page in pts
    ''' (topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance)
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_getDimensions_CoverPage_Prt() As Collection
        Dim lstOfDimensions As New Collection()
        Dim objRptMgr As New cReport()
        '
        'topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance
        '
        Call Me.initPageSettings(lstOfDimensions, 162.0#, 56.0#, 56.0#, 56.0#, 0.0#, 22.7, 8.2)              'Long and Short Report Portrait Page
        '
        Return lstOfDimensions
        '
    End Function

    '
    ''' <summary>
    ''' This method will return the dimensions of a standard report Cover Landscape Page in pts
    ''' (topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance)
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_getDimensions_CoverPage_Lnd() As Collection
        Dim lstOfDimensions As New Collection()
        Dim leftedge As Single
        '
        leftedge = Me.glb_math_MillimetersToPoints(Me._glb_header_leftEdge)
        '
        'Call Me.initPageSettings(lstOfDimensions, 56.7#, 155.95, 56.7#, 42.55, 0.0#, 22.7, 7.95)
        'Call Me.initPageSettings(lstOfDimensions, 56.7#, leftedge, 56.7#, 42.55, 0.0#, 22.7, 7.95)
        Call Me.initPageSettings(lstOfDimensions, 128.0#, 53.4, 48.0#, 422.0#, 0.0#, 22.7, 7.95)

        '
        Return lstOfDimensions
        '
    End Function
    '

    '
    Public Function glb_math_MillimetersToPoints(measurementInmm As Single)
        glb_math_MillimetersToPoints = 72.0 * (measurementInmm / 25.4)
    End Function
    '

    '
    ''' <summary>
    ''' This method will return the dimensions of a Landscape report Landscape Page in pts
    ''' (topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance)
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_getDimensions_Landscape_LndRpt_followerPage() As Collection
        Dim lstOfDimensions As New Collection()
        '
        'Call Me.initPageSettings(lstOfDimensions, 56.0#, 156.0, 56.0#, 42.55, 0.0#, 22.7, 7.95)
        Call Me.initPageSettings(lstOfDimensions, 56.0#, 156.0, 56.0#, 50.0#, 0.0#, 22.7, 7.95)
        '
        Return lstOfDimensions
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will return the dimensions of a Landscape report Chapter Page in pts
    ''' (topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance)
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_getDimensions_Landscape_LndRpt_ChptPage() As Collection
        Dim lstOfDimensions As New Collection()
        '
        'Call Me.initPageSettings(lstOfDimensions, 72.0#, 72.0, 56.7#, 42.55, 0.0#, 22.7, 7.95)
        Call Me.initPageSettings(lstOfDimensions, 72.0#, 72.0, 56.7#, 50.0#, 0.0#, 22.7, 7.95)
        '
        Return lstOfDimensions
        '
    End Function
    '
    '
    '
    ''' <summary>
    ''' This method will return the dimensions of a standard letter in pts
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_getDimensions_Letter() As Collection
        Dim lstOfDimensions As New Collection()
        '
        'Call Me.initPageSettings(lstOfDimensions, 80.0#, 55.8, 84.0, 42.55#, 0.0#, 23.4, 3.8)
        Call Me.initPageSettings(lstOfDimensions, 80.0#, 56.0, 84.0, 42.55#, 0.0#, 23.4, 3.8)
        '
        Return lstOfDimensions
        '
    End Function
    '
    ''' <summary>
    ''' This is a generic method that sets the dimensions of the specified sect according to the
    ''' provided lstOfDimensions
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="lstOfDimensions"></param>
    Public Sub glb_setDimensions(ByRef sect As Word.Section, ByRef lstOfDimensions As Collection)
        sect.PageSetup.TopMargin = CSng(lstOfDimensions("topMargin"))               'top
        sect.PageSetup.LeftMargin = CSng(lstOfDimensions("leftMargin"))             'left
        sect.PageSetup.BottomMargin = CSng(lstOfDimensions("bottomMargin"))         'bottom
        sect.PageSetup.RightMargin = CSng(lstOfDimensions("rightMargin"))           'right
        sect.PageSetup.Gutter = CSng(lstOfDimensions("gutter"))                     'gutter
        sect.PageSetup.HeaderDistance = CSng(lstOfDimensions("headerDistance"))     'header
        sect.PageSetup.FooterDistance = CSng(lstOfDimensions("footerDistance"))     'footer
        '
    End Sub

    Sub initPageSettings(ByRef lstOfItems As Collection, topMargin As Single, leftMargin As Single,
                                    bottomMargin As Single, rightMargin As Single, gutter As Single,
                                    headerDistance As Single, footerDistance As Single)
        '
        lstOfItems.Clear()
        '
        Call lstOfItems.Add(topMargin, "topMargin")
        Call lstOfItems.Add(leftMargin, "leftMargin")
        Call lstOfItems.Add(bottomMargin, "bottomMargin")
        Call lstOfItems.Add(rightMargin, "rightMargin")
        Call lstOfItems.Add(gutter, "gutter")
        Call lstOfItems.Add(headerDistance, "headerDistance")
        Call lstOfItems.Add(footerDistance, "footerDistance")
        '
    End Sub
    '
    '
    Public Sub glb_Set_LetterPageDimensions(ByRef sect As Section)

        Call glb_setDimensions(sect, Me.glb_getDimensions_Letter)
    End Sub
    '
    '
    ''' <summary>
    ''' This method will return the dimensions of the standard TOC Page in pts
    ''' (topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance)
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_getDimensions_toc_Prt() As Collection
        Dim lstOfDimensions As New Collection()
        '
        '(topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance)
        'Call Me.initPageSettings(lstOfDimensions, 56.7#, 155.95, 56.7#, 42.55, 0.0#, 22.7, 7.95)                'Portrait Report 15mm right margin
        '
        '*** Test settings
        'Call Me.initPageSettings(lstOfDimensions, 56.0#, 120, 56.0#, 30, 0.0#, 22.7, 7.95)                'Portrait Report 15mm right margin
        'Call Me.initPageSettings(lstOfDimensions, 56.0#, 120, 76.0#, 50.0, 0.0#, 22.7, 7.95)                'As per sepc 20231204
        Call Me.initPageSettings(lstOfDimensions, 56.0#, 120, 54.0#, 50.0, 0.0#, 22.7, 7.95)                'As per sepc 20231204

        '
        Return lstOfDimensions
        '
        '
    End Function
    '
    ''' <summary>
    ''' This method will return the dimensions of the standard TOC Page in pts
    ''' (topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance)
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_getDimensions_toc_Lnd() As Collection
        Dim lstOfDimensions As New Collection()
        Dim objRptMgr As New cReport()
        '
        'topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance
        '
        lstOfDimensions = Me.glb_getDimensions_Std_Lnd()
        '
        Return lstOfDimensions
        '
    End Function

    '
#End Region
    '
#Region "Document Items"
    ''' <summary>
    ''' This method will return the centre point between the two margins of a page/section
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function glb_get_tabCenterPos(ByRef sect As Word.Section) As Single
        Dim tabCentre As Single
        '
        tabCentre = sect.PageSetup.LeftMargin / 2 + sect.PageSetup.PageWidth / 2 - sect.PageSetup.RightMargin / 2
        '
        Return tabCentre
    End Function
    '
    '
    '
    ''' <summary>
    ''' This method will return the Style that has the name 'strStyleName'. If there
    ''' is an error it will return the Normal style
    ''' </summary>
    ''' <param name="strStyleName"></param>
    ''' <returns></returns>
    Public Function glb_get_wrdStyle(strStyleName As String) As Word.Style
        Dim rslt As Word.Style
        '
        rslt = Nothing
        '
        rslt = glb_get_wrdActiveDoc.Styles(strStyleName)
        '
        'Try
        'rslt = glb_get_wrdDoc.Styles(strStyleName)
        'Catch ex As Exception
        'rslt = glb_get_wrdDoc.Styles("Normal")
        'End Try
        '
        Return rslt
    End Function
    '
    Public Sub glb_cursors_setToWait()
        glb_get_wrdApp.System.Cursor = WdCursorType.wdCursorWait
    End Sub
    '
    Public Sub glb_cursors_setToNormal()
        glb_get_wrdApp.System.Cursor = WdCursorType.wdCursorNormal
    End Sub

    Public Function glb_doc_isTemplate() As Boolean
        Return False
    End Function
    '
    ''' <summary>
    ''' This method will determine if the template attached to myDoc is the standard
    ''' ACIL Allen template. It does by testing 'If tmpl.FullName Like "*AA GeneralReport.dotx" Then'
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function glb_doc_hasAAStdTemplate(ByRef myDoc As Word.Document) As Boolean
        Dim tmpl As Word.Template
        Dim rslt As Boolean
        '
        rslt = False
        tmpl = myDoc.AttachedTemplate
        '
        If tmpl.FullName Like "*AA GeneralReport.dotx" Then rslt = True

        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' This method will determine if the document myDoc is a standard ACIL Allen
    ''' document, and as a consequence able to respond reliably to the ribbon functions.
    ''' At the moment it checks to see if the attached template is the standard AA template.
    ''' Later we might check for some other 'fingerprint'. Maybe styles??
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function glb_doc_isAAStdDoc(ByRef myDoc As Word.Document) As Boolean
        Dim rslt As Boolean
        '
        'rslt = Me.glb_doc_hasAAStdTemplate(myDoc)
        rslt = False
        '
        If glb_style_Exists(myDoc, "tag_chapterBanner") And Me.glb_doc_hasAAStdTemplate(myDoc) Then
            rslt = True
        End If
        'rslt = glb_style_Exists(myDoc, "tag_chapterBanner")         'Use this to test for AA Std document.. Could alsu use tag_aa_RptTestStyle_#?_00
        'rslt = glb_style_Exists(myDoc, "tag_aa_RptTestStyle_#?_00")         'Use this to test for AA Std document.. Could alsu use tag_aa_RptTestStyle_#?_00
        ' 
        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' This method will return true if a specific style with name stylename exists
    ''' in myDoc
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="styleName"></param>
    ''' <returns></returns>
    Public Function glb_style_Exists(myDoc As Word.Document, styleName As String) As Boolean
        '
        Try
            Dim myStyle As Word.Style = myDoc.Styles.Item(styleName)
            Return Not IsNothing(myStyle)
        Catch ex As Exception
            Return False
        End Try
        '
    End Function
    '
    ''' <summary>
    ''' This method is used to detremine if a document is empty. It checks the main body,
    ''' headers and footers for content. If any of these contain text, it returns False.
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function glb_doc_isEmptyAndNotSaved(ByRef myDoc As Word.Document) As Boolean
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim rslt As Boolean = True
        Dim hasNoText As Boolean = False
        Dim isNotSaved As Boolean = False
        Dim bodyText As String = myDoc.Content.Text.Trim()
        Dim hfHasNoContent As Boolean = False
        '
        If myDoc.Path = "" Then isNotSaved = True
        If bodyText.Length = 0 Then hasNoText = True
        hfHasNoContent = Not objHFMgr.hf_hfs_haveContent(myDoc)
        '
        rslt = isNotSaved And hasNoText And hfHasNoContent
        '
        Return rslt

    End Function

    '
    ''' <summary>
    ''' Ths method will check the doc type (AA Std or not) of the current Active Document and 
    ''' activate 'Pages and Sections' tab if its a ACIL Allen document, or the 'Home' tab if it is not
    ''' </summary>
    ''' <param name="strTabId"></param>
    ''' <returns></returns>
    Public Function glb_doc_checkDocType_ActivateTab(Optional strTabId As String = "tab_aa_PagesAndSections") As Word.Document
        Dim myDoc As Word.Document
        Dim objCtrls As New cControlsMgr()
        'Dim rbn As rbn_aa_Addin00
        '
        myDoc = glb_get_wrdActiveDoc()
        '
        '*** leave all tabs on for testing on normal documents
        '
        'objCtrls.ctrl_tabSet_Visibility("all")
        'Globals.Ribbons.rbn_Styles.RibbonUI.ActivateTab(strTabId)
        'GoTo finis
        '***
        '
        If Me.glb_doc_isAAStdDoc(Me.glb_get_wrdActiveDoc) Then
            objCtrls.ctrl_tabSet_Visibility("all")
            Globals.Ribbons.rbn_Styles.RibbonUI.ActivateTab(strTabId)
            'Globals.Ribbons.rbn_Styles.tab_aa_PagesAndSections
        Else
            objCtrls.ctrl_tabSet_Visibility("all", False)
            Globals.Ribbons.rbn_Styles.RibbonUI.ActivateTab(objCtrls._strTabId_AAHome)

            'Globals.Ribbons.rbn_Styles.RibbonUI.ActivateTabMso("TabHome")
            'Globals.Ribbons.rbn_Styles.RibbonUI.ActivateTabMso("TabHome")
            'Globals.Ribbons.rbn_Styles.RibbonUI.ActivateTabMso("TabHome")
            '
        End If
        '
finis:
        Return myDoc
    End Function
    '
    ''' <summary>
    ''' This method will de-activate the ACIL Allen tabs
    ''' </summary>
    Public Sub glb_deactivate_AATabs()
        Dim objCtrls As New cControlsMgr()
        '
        objCtrls.ctrl_tabSet_Visibility("all", False)
        '
    End Sub


    ''' <summary>
    ''' This method will return the Active Document
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_get_wrdActiveDoc() As Word.Document
        Dim myDoc As Word.Document
        '
        'myDoc = Globals.ThisDocument.Application.ActiveDocument
        myDoc = Globals.ThisAddIn.Application.ActiveDocument
        '
        Return myDoc
    End Function
    '
    ''' <summary>
    ''' This method will set the field shading of the ActiveDocument depending on on the
    ''' value of strShadingMode. It can take on the values; 'always', 'never' or 'whenSelected'
    ''' </summary>
    ''' <param name="strShadingMode"></param>
    Public Sub glb_set_fieldShading(strShadingMode As String)
        '
        Select Case strShadingMode
            Case "always"
                glb_get_wrdActiveDoc.ActiveWindow.View.FieldShading = WdFieldShading.wdFieldShadingAlways
            Case "never"
                glb_get_wrdActiveDoc.ActiveWindow.View.FieldShading = WdFieldShading.wdFieldShadingNever
            Case "whenSelected"
                glb_get_wrdActiveDoc.ActiveWindow.View.FieldShading = WdFieldShading.wdFieldShadingWhenSelected
        End Select
        '
    End Sub

    '
    ''' <summary>
    ''' This method will return the Application
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_get_wrdApp() As Word.Application
        Dim myApp As Word.Application
        '
        'myApp = Globals.ThisDocument.Application
        myApp = Globals.ThisAddIn.Application
        '
        Return myApp
    End Function
    '
    ''' <summary>
    ''' This method will return 'prt' or 'lnd' depending on the orientation of the section
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function glb_sect_getOrientation(ByRef sect As Word.Section) As String
        Dim strResult As String
        '
        strResult = "prt"
        '
        If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then strResult = "lnd"

        Return strResult
    End Function
    '
    '
    ''' <summary>
    ''' This method will return the first cell in the Application Selection. If
    ''' no cell, then it will rturn nothing
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_get_wrdSelCell() As Word.Cell
        Dim mySel As Word.Selection
        Dim drCell As Word.Cell
        '
        drCell = Nothing
        '
        Try
            mySel = Me.glb_get_wrdSel()
            If mySel.Cells.Count <> 0 Then
                drCell = mySel.Cells.Item(1)
            Else
                drCell = Nothing
            End If
            'mySel = Globals.ThisAddin.Application.Selection
        Catch ex As Exception
            drCell = Nothing
        End Try
        '
        Return drCell
    End Function
    '

    '
    ''' <summary>
    ''' This method will return the Application Selection
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_get_wrdSel() As Word.Selection
        Dim mySel As Word.Selection
        '
        'mySel = Globals.ThisDocument.Application.Selection
        mySel = Globals.ThisAddIn.Application.Selection
        '
        Return mySel
    End Function
    '
    ''' <summary>
    ''' This method looks for any tables in the Selection range. If there are any
    ''' at all it will return true. Typically this approach will also return true 
    ''' if the selection is in a paragrapgh just below a Table
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_selection_IsInTable() As Boolean
        Dim rslt As Boolean
        'Dim rng As Word.Range
        '
        rslt = False
        '
        rslt = glb_get_wrdSel().Information(WdInformation.wdWithInTable)
        '
        'rng = Me.glb_get_wrdSel.Range
        'If rng.Tables.Count <> 0 Then rslt = True
        '
        Return rslt
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method assumes text in the cell drCell. It will set the seelction to that text
    ''' </summary>
    ''' <param name="drCell"></param>
    Public Sub glb_selection_toCellText(ByRef drCell As Word.Cell)
        Dim rng As Word.Range
        '
        Try
            rng = drCell.Range
            rng.MoveEnd(WdUnits.wdCharacter, -1)
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Select()
        Catch ex As Exception

        End Try

    End Sub
    '
    ''' <summary>
    ''' This method will return the section index of the section that contains
    ''' the start of the current selection
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_get_wrdSelSectIdx() As Integer
        Dim rng As Word.Range
        Dim idx As Integer
        '
        rng = Me.glb_get_wrdApp().Selection.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        idx = rng.Information(WdInformation.wdActiveEndSectionNumber)
        '
        Return idx
    End Function
    '
    ''' <summary>
    ''' This method will return the Application Selection (collapsed to start) range
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_get_wrdSelRng() As Word.Range
        Dim rng As Word.Range
        '

        rng = Me.glb_get_wrdApp().Selection.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        Return rng
    End Function
    '
    ''' <summary>
    ''' This method will return the entire range of the Selection
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_get_wrdSelRngAll() As Word.Range
        Dim rng As Word.Range
        '
        rng = Me.glb_get_wrdApp().Selection.Range
        '
        Return rng
    End Function
    '   
    ''' <summary>
    ''' This method will return the first table in the current selection. If the selection
    ''' does not contain a Table it will return nothing
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_get_wrdSelTbl() As Word.Table
        Dim tbl As Word.Table
        Dim rng As Word.Range
        '
        tbl = Nothing
        rng = Me.glb_get_wrdApp().Selection.Range
        '
        If rng.Tables.Count <> 0 Then
            tbl = rng.Tables.Item(1)
        Else
            tbl = Nothing
        End If
        'tbl = Me.glb_get_wrdSelTbl2()
        '
        Return tbl
    End Function
    '
    '   
    ''' <summary>
    ''' This method will return the first table in the current selection. If the selection
    ''' does not contain a Table it will return nothing
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_get_wrdSelTbl2() As Word.Table
        Dim tbl As Word.Table
        '
        tbl = Nothing
        '
        Try
            If Me.glb_get_wrdApp.Selection.Information(WdInformation.wdWithInTable) Then
                tbl = Me.glb_get_wrdApp.Selection.Tables(1)
                'Me.tb
                '
            Else
                tbl = Nothing
            End If
            '
        Catch ex As Exception
            tbl = Nothing
        End Try
        '
        Return tbl
    End Function
    '

    '
    '
    ''' <summary>
    ''' This method will return the section that contains the beginning of the
    ''' current selection
    ''' </summary>
    ''' <returns></returns>
    Public Function glb_get_wrdSect() As Word.Section
        Dim sect As Word.Section
        Dim rng As Word.Range
        '
        rng = Me.glb_get_wrdSelRng()
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        sect = rng.Sections.Item(1)
        '
        Return sect
    End Function
    '
    ''' <summary>
    ''' This method will set the selection so that there are 'numParas' between the table and the selection.
    ''' A value of '0' for 'numParas' means that the selection (and range rng) are directly below the table.
    ''' Typically this would be used when you want to ensure a specific spacing between a table such as a banner
    ''' or placeholder and the point where you may want to insert a new table. The return value is the range of 
    ''' the selection.... This method adds empty paras
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function glb_set_wrdSel(ByRef tbl As Word.Table, Optional numParas As Integer = 0) As Word.Range
        Dim rng As Word.Range
        Dim i As Integer
        '
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        'rng.Move(WdUnits.wdParagraph, 1)
        '
        If numParas <= 0 Then GoTo finis
        '
        For i = 1 To numParas
            rng.Paragraphs.Add(rng)
        Next
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
finis:
        rng.Select()
        '
        Return rng
    End Function
#End Region
    '
    ''' <summary>
    ''' This method will add a section break at the range rng (typically collapsed). It will do so according to the
    ''' break type strBreakType which can be 'newPage', 'oddPage', or 'evenPage'. It will return the new section
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="strBreakType"></param>
    ''' <returns></returns>
    Public Function glb_add_sectionBreak(ByRef rng As Word.Range, Optional strBreakType As String = "newPage") As Word.Section
        Dim myDoc As Word.Document
        Dim sect As Word.Section
        '
        sect = rng.Sections.First
        myDoc = sect.Range.Document
        '
        Select Case strBreakType
            Case "newPage"
                sect = myDoc.Sections.Add(rng, WdSectionStart.wdSectionNewPage)
            Case "oddPage"
                sect = myDoc.Sections.Add(rng, WdSectionStart.wdSectionOddPage)
            Case "evenPage"
                sect = myDoc.Sections.Add(rng, WdSectionStart.wdSectionEvenPage)
            Case Else
                sect = myDoc.Sections.Add(rng, WdSectionStart.wdSectionNewPage)
        End Select
        '
        Return sect
    End Function
End Class
