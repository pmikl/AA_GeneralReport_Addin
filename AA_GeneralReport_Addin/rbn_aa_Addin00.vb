Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports System.Drawing
Imports stdole

Public Class rbn_aa_Addin00

    Private Sub rbn_aa_Addin00_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        MsgBox("Ribbon is loaded")
    End Sub
    '
#Region "Support routines"
    Public Sub rbn_response_toDownLoad(strFullFileName As String)
        Dim strMsg As String
        Dim myFileInfo As New System.IO.FileInfo(strFullFileName)
        '
        Select Case strFullFileName
            Case "cancel"
                MsgBox("Download has been cancelled by the user")
            Case ""
                MsgBox("Error in file download..")
            Case Else
                strMsg = "Download complete. The file is located at" + vbCrLf + strFullFileName + vbCrLf + vbCrLf + "Do you want the file opened"
                'MsgBox("Download complete. The file is located at" + vbCrLf + strFilePath)
                If myFileInfo.Extension = ".docx" Or myFileInfo.Extension = ".dotx" Then
                    If MsgBox(strMsg, vbYesNo + vbDefaultButton2, "Resource ownload status") = MsgBoxResult.Yes Then
                        Globals.ThisAddIn.Application.Documents.Open(strFullFileName)
                    End If
                Else
                    MsgBox("The file with extension " + myFileInfo.Extension + " cannot be opened by Word")
                End If
        End Select
        '
    End Sub
#End Region


    Private Sub gal_CoverPages_Prt_ItemsLoading2(sender As Object, e As RibbonControlEventArgs)
        Dim gallery As RibbonGallery = CType(sender, RibbonGallery)
        Dim myImage As Image
        'gallery.Items.Clear()

        ' Load image from resources
        'Dim resizedImage As New Bitmap(myImage, New Size(48, 48)) ' Resize to 48x48
        'Dim pictureDisp As stdole.IPictureDisp = RibbonConverter.ToIPictureDisp(resizedImage)
        '
        myImage = My.Resources.CP_TG_filledPattern__lnd_small
        gallery.Items(0).Label = "test 0"
        gallery.Items(0).Image = myImage
        'gallery.Items(0).Image.
        'gallery.Items(0).Image.Size.Height(

        'myImage = My.Resources.Cp_TG_emptyPattern_lnd_small
        'gallery.Items(1).Label = "test 1"
        ' gallery.Items(1).Image = myImage
        '
        'myImage = My.Resources.Cp_TG_picturePattern_lnd_small
        'gallery.Items(2).Label = "test 2"
        'gallery.Items(2).Image = myImage

        'galleryItem1.Label = "Style 1"
        'galleryItem1.Image = RibbonConverter.ToIPictureDisp(myImage)

        ' Add it back to the gallery
        'gallery.Items.Add(galleryItem1)
        'gallery.ShowImage = True
    End Sub
    '
    Private Sub gal_CoverPages_Click(sender As Object, e As RibbonControlEventArgs) Handles gal_CoverPages.Click
        Dim objGlobals As New cGlobals()
        Dim gallery As RibbonGallery = CType(sender, RibbonGallery)
        Dim objCpMgr As cCoverPageMgr
        Dim objBBlockMgr As cBBlocksHandler
        Dim objRptMgr As New cReport()
        Dim objHFMgr As cHeaderFooterMgr
        Dim objMsgMgr As New cMessageManager()
        Dim objWorkAround As New cWorkArounds()
        Dim sect As Word.Section
        Dim myDoc As Word.Document
        Dim tokens() As String
        Dim strRptMode, selectedId As String
        '
        'Me.ribbon.InvalidateControl(control.Id)
        'Me.ribbon.InvalidateControl("grp_PgNumMgmnt_chBxNumFormat_2part")

        '
        sect = Nothing
        myDoc = objGlobals.glb_get_wrdActiveDoc()
        strRptMode = objRptMgr.Rpt_Mode_Get()
        '
        'objGlobals.glb_screen_update(False)

        '
        'We set the gallery item ids to reflect the cover page type (see _galleryGetItemId)
        selectedId = gallery.SelectedItem.Tag.ToString()
        tokens = Split(selectedId, "-")
        '
        Select Case strRptMode
            Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                Try
                    Select Case e.Control.Id
                        Case "gal_CoverPages"
                            objCpMgr = New cCoverPageMgr()
                            objBBlockMgr = New cBBlocksHandler()
                            objHFMgr = New cHeaderFooterMgr()
                            '
                            sect = objCpMgr.cp_Insert_CoverPage(myDoc, tokens(1))
                            objGlobals.glb_view_setToPrintLayout()
                            '
                            objWorkAround.wrk_fix_forCursorRace()
                            '
                            objCpMgr.cp_set_SelectionToTitle(sect)
                            '
                    End Select
                Catch ex As Exception
                    objGlobals.glb_view_setToPrintLayout()
                    MsgBox("Insertion Error")
                End Try
                '
            Case objRptMgr.rpt_isBrief
                MsgBox(objMsgMgr.msgMgr_msg_notAvailableInBrief())
                '
        End Select
        '
        '
        objGlobals.glb_screen_update(True)
        '
        objWorkAround.wrk_fix_forCursorRace()
        '
        objGlobals.glb_cursors_setToNormal()
    End Sub

    Private Sub gal_CoverPages_ItemsLoading(sender As Object, e As RibbonControlEventArgs) Handles gal_CoverPages.ItemsLoading
        Dim objGlobals As New cGlobals()
        Dim gallery As RibbonGallery = CType(sender, RibbonGallery)
        Dim sect As Word.Section
        Dim myImage As Image
        '
        sect = objGlobals.glb_get_wrdSect()

        If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then
            myImage = My.Resources.CP_TG_filledPattern__lnd_small
            gallery.Items(0).Label = "Lnd 0"
            gallery.Items(0).Image = myImage

            myImage = My.Resources.Cp_TG_emptyPattern_lnd_small
            gallery.Items(1).Label = "Lnd 1"
            gallery.Items(1).Image = myImage
            '
            myImage = My.Resources.Cp_TG_picturePattern_lnd_small
            gallery.Items(2).Label = "Lnd 2"
            gallery.Items(2).Image = myImage
            '
        Else
            myImage = My.Resources.CP_TG_filledPattern_small
            gallery.Items(0).Label = "Lnd 0"
            gallery.Items(0).Image = myImage

            myImage = My.Resources.Cp_TG_emptyPattern_small
            gallery.Items(1).Label = "Lnd 1"
            gallery.Items(1).Image = myImage
            '
            myImage = My.Resources.Cp_TG_picturePattern_small
            gallery.Items(2).Label = "Lnd 2"
            gallery.Items(2).Image = myImage
            '
        End If
        '
        gallery.ShowImage = True
        '
    End Sub


    'PIF_tab00_grpAA00

    Private Sub PIF_Styl_grpApp_Click(sender As Object, e As RibbonControlEventArgs) Handles grpStylesApp_StyleSet.Click, grpStylesApp_Heading5_App.Click, grpStylesApp_Heading4_App.Click, grpStylesApp_Heading3_App.Click, grpStylesApp_Heading2_App.Click, grpStylesApp_Heading1_App.Click
        Dim objStyles As New cStylesManager()
        Dim objWrkAround As New cWorkArounds()
        Dim objSectMgr As New cSectionMgr()
        Dim strStyleName As String
        '
        objSectMgr.objGlobals.glb_screen_update(False)
        '
        strStyleName = "Normal"
        '
        Select Case e.Control.Id
            Case "grpStylesApp_StyleSet"
                Call objStyles.styles_format_styleSetAP()
            Case "grpStylesApp_Heading1_App"
                If objSectMgr.objGlobals._glb_doApp_as_HeadingAP Then
                    strStyleName = "Heading 1 (AP)"
                Else
                    strStyleName = "Heading 6"
                End If
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesApp_Heading2_App"
                If objSectMgr.objGlobals._glb_doApp_as_HeadingAP Then
                    strStyleName = "Heading 2 (AP)"
                Else
                    strStyleName = "Heading 7"
                End If
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesApp_Heading3_App"
                If objSectMgr.objGlobals._glb_doApp_as_HeadingAP Then
                    strStyleName = "Heading 3 (AP)"
                Else
                    strStyleName = "Heading 8"
                End If
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesApp_Heading4_App"
                If objSectMgr.objGlobals._glb_doApp_as_HeadingAP Then
                    strStyleName = "Heading 4 (AP)"
                Else
                    strStyleName = "Heading 9"
                End If
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesApp_Heading5_App"
                strStyleName = "Heading 5 (AP)"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesApp_Heading6_App"
                strStyleName = "Heading 6 (AP)"
                objStyles.applyStyleToSelection(strStyleName)
        End Select
        '
        objSectMgr.objGlobals.glb_screen_update(True)
        '
        objWrkAround.wrk_fix_forCursorRace()
        '
    End Sub

    Private Sub PIF_tabStyles_Click(sender As Object, e As RibbonControlEventArgs) Handles xbtn_mnuThemes_PGSToggle.Click, xbtn_mnuThemes_ActivateTabPGS.Click, xbtn__mnuThemes_set_AAThemeAndStyles.Click, xbtn__mnuThemes_set_AATheme.Click, tabThms_mnu_Set_btn_PGSToggle.Click, tabThms_mnu_Set_btn_ActivateTabPGS.Click, tabThms_mnu_Set_btn__applyAAThemeStylesTemplate.Click, tabThms_mnu_Set_btn_applyAATheme.Click, tabThms_mnu_Set_btn_attachNormalTemplate.Click, tabThms_btn_resetStylesForRptPrt.Click, tabThms_btn_resetStylesForRptLnd.Click, tabThms_btn_resetStylesForRptBrf.Click, tabThms_mnu_Set_btn_getAttachedTemplate.Click, tabThms_mnu_Set_btn_attachAATemplate.Click, tabStyles_btn_resetStylesForRptPrt.Click, tabStyles_btn_resetStylesForRptLnd.Click, tabStyles_btn_resetStylesForRptBrf.Click
        Dim objThmMgr As New cThemeMgr()
        Dim objGlobals As New cGlobals()
        Dim objRptMgr As New cReport()
        Dim objStylesMgr As New cStylesManager()
        Dim objMsgMgr As New cMessageManager()
        Dim objTOCMgr As New cTOCMgr()
        Dim strRptMode As String
        Dim tmpl As Word.Template
        'Globals.Ribbons.rbn_Styles.RibbonUI.ActivateTab("tab_aa_PagesAndSections")
        '
        strRptMode = objRptMgr.Rpt_Mode_Get()
        '
        Select Case e.Control.Id
            Case "tabThms_mnu_Set_btn_applyAATheme"
                objThmMgr.thm_Set_ThemeToAAStd_fromFile(objGlobals.glb_get_wrdActiveDoc)
                objGlobals.glb_doc_checkDocType_ActivateTab(objGlobals._strTabId_AAHome)
                '
                MsgBox("The standard ACIL Allen theme has been applied to the current document")
                '
            Case "tabThms_mnu_Set_btn__applyAAThemeStylesTemplate"
                If objMsgMgr.msgMgr_dlg_themesStylesAndTemplateWarning() Then
                    objThmMgr.thm_Set_ThemeToAAStd_fromFile(objGlobals.glb_get_wrdActiveDoc)
                    'objStylesMgr.styl_build_AAStyles(objGlobals.glb_get_wrdActiveDoc)
                    objRptMgr.Rpt_Styles_resetStyles_fromTemplate(True)
                    '
                    'objGlobals.glb_doc_checkDocType_ActivateTab(objGlobals._strTabId_ThemesAndTest)
                    objStylesMgr.glb_doc_checkDocType_ActivateTab(objGlobals._strTabId_PagesAndSections)

                    '
                Else
                    MsgBox("The function has been cancelled by the user")
                End If
                '
            Case "tabThms_mnu_Set_btn_attachNormalTemplate"
                objGlobals.glb_get_wrdActiveDoc.AttachedTemplate = "Normal"
                objGlobals.glb_doc_checkDocType_ActivateTab(objGlobals._strTabId_AAHome)
                '
            Case "tabThms_mnu_Set_btn_attachAATemplate"
                objGlobals.glb_get_wrdActiveDoc.AttachedTemplate = objGlobals.glb_getTmpl_FullName()
                objStylesMgr.glb_doc_checkDocType_ActivateTab(objGlobals._strTabId_PagesAndSections)
                '
            Case "tabThms_mnu_Set_btn_getAttachedTemplate"
                tmpl = objGlobals.glb_get_wrdActiveDoc.AttachedTemplate
                MsgBox("The attached template is " + tmpl.FullName)
                objGlobals.glb_doc_checkDocType_ActivateTab(objGlobals._strTabId_AAHome)

            Case "tabThms_mnu_Set_btn_ActivateTabPGS"
                objGlobals.ctrl_tab_Activate(objGlobals._strTabId_PagesAndSections)
                MsgBox("Activate PGS")

            Case "tabThms_mnu_Set_btn_PGSToggle"
                objGlobals.ctrl_tabToggle_Visibility(objGlobals._strTabId_PagesAndSections)
                '
            Case "tabThms_btn_resetStylesForRptPrt", "tabStyles_btn_resetStylesForRptPrt"
                objThmMgr.thm_Set_ThemeToAAStd_fromFile(objGlobals.glb_get_wrdActiveDoc)
                '
                objRptMgr.Rpt_Styles_resetStyles_fromTemplate(True)
                objStylesMgr.style_extend_TemplateStyles()
                '
                Select Case e.Control.Id
                    Case "tabThms_btn_resetStylesForRptPrt"
                        objGlobals.glb_doc_checkDocType_ActivateTab(objGlobals._strTabId_AAHome)
                    Case "tabStyles_btn_resetStylesForRptPrt"
                        objGlobals.glb_doc_checkDocType_ActivateTab(objGlobals._strTabId_Styles)
                End Select
                '
            Case "tabThms_btn_resetStylesForRptLnd", "tabStyles_btn_resetStylesForRptLnd"
                objThmMgr.thm_Set_ThemeToAAStd_fromFile(objGlobals.glb_get_wrdActiveDoc)
                '
                objRptMgr.Rpt_Styles_resetStyles_fromTemplate(True)
                objStylesMgr.style_extend_TemplateStyles()                                                                      'Refresh the styles
                objRptMgr.Rpt_styles_Upgrade_for_ReportType(objGlobals.glb_get_wrdActiveDoc, objRptMgr.rpt_isLnd)        'Upgrade/chnage depending on report mode
                objTOCMgr.TOC_Styles_AdjustForReportMode(objRptMgr.rpt_isLnd)                                                     'Force the style for the Brief
                '
                Select Case e.Control.Id
                    Case "tabThms_btn_resetStylesForRptLnd"
                        objGlobals.glb_doc_checkDocType_ActivateTab(objGlobals._strTabId_AAHome)
                    Case "tabStyles_btn_resetStylesForRptLnd"
                        objGlobals.glb_doc_checkDocType_ActivateTab(objGlobals._strTabId_Styles)
                End Select
                '
            Case "tabThms_btn_resetStylesForRptBrf", "tabStyles_btn_resetStylesForRptBrf"
                objThmMgr.thm_Set_ThemeToAAStd_fromFile(objGlobals.glb_get_wrdActiveDoc)
                '
                objRptMgr.Rpt_Styles_resetStyles_fromTemplate(True)
                objStylesMgr.style_extend_TemplateStyles()                                                                          'Refresh the styles
                objRptMgr.Rpt_styles_Upgrade_for_ReportType(objGlobals.glb_get_wrdActiveDoc, objRptMgr.rpt_isBrief)                 'Force the styles for the Brief
                objTOCMgr.TOC_Styles_AdjustForReportMode(objRptMgr.rpt_isBrief)                                                     'Force the style for the Brief
                '
                Select Case e.Control.Id
                    Case "tabThms_btn_resetStylesForRptBrf"
                        objGlobals.glb_doc_checkDocType_ActivateTab(objGlobals._strTabId_AAHome)
                    Case "tabStyles_btn_resetStylesForRptBrf"
                        objGlobals.glb_doc_checkDocType_ActivateTab(objGlobals._strTabId_Styles)
                End Select
                '
        End Select
    End Sub

    Private Sub PIF_tabPGS_grpRpt_Click(sender As Object, e As RibbonControlEventArgs)
        Dim objThmMgr As New cThemeMgr()
        Dim objGlobals As New cGlobals()
        Dim objRptMgr As New cReport()


        'Globals.Ribbons.rbn_Styles.RibbonUI.ActivateTab("tab_aa_PagesAndSections")
        '
        Select Case e.Control.Id
            Case "grpRpt_btn_GlossaryAndAbbreviations"
                '
            Case "grpReport_btn_newDivider_Chpt"
                '
            Case "btn_mnuThemes_ActivateTabPGS"
                objGlobals.ctrl_tab_Activate(objGlobals._strTabId_PagesAndSections)
                MsgBox("Activate PGS")

            Case "grpReport_btn_ToggleView"
                '

        End Select

    End Sub
    '
    Private Sub PIF_Styl_grpResetStyles_Click(sender As Object, e As RibbonControlEventArgs) Handles grpStylesTools_to_PrintDefault.Click, grpStylesTools_resetCaptions.Click, grpStylesTools_to_DisplayDefault.Click
        Dim objStyles As New cStylesManager()
        Dim objCaptionsMgr As New cCaptionManager()
        Dim objRptMgr As New cReport()
        Dim objGlobals As New cGlobals()
        Dim objPrint As cPrintAndDisplayServices
        Dim rng As Word.Range

        Dim strStyleName As String
        '
        objGlobals.glb_get_wrdApp.ScreenUpdating = False
        strStyleName = "Normal"
        '
        Select Case e.Control.Id
            Case "grpStylesTools_to_PrintDefault"
                objPrint = New cPrintAndDisplayServices()
                objPrint.colour_display_ToDesignView(objGlobals.glb_get_wrdActiveDoc())
                '
            Case "grpStylesTools_to_DisplayDefault"
                objPrint = New cPrintAndDisplayServices()
                objPrint.colour_display_ToEasyView(objGlobals.glb_get_wrdActiveDoc())
            Case "grpStylesTools_resetStyle"
                rng = objGlobals.glb_get_wrdSel.Range
                Try
                    objRptMgr.Rpt_Styles_resetStyles_fromTemplate()
                    MsgBox("The report styles have been reset from the official AAC template")
                Catch ex As Exception

                End Try
                '                
                rng.Select()
                '
            Case "grpStylesTools_get_TemplateStyles"
                '
                If Not objStyles.style_copy_StylesFromTemplate(objGlobals.glb_get_wrdActiveDoc) Then
                    MsgBox("Styles 'flow' update from the template has failed")
                End If

            Case "grpStylesTools_resetCaptions"
                MsgBox("This function will rebuild the Custom Captions (e.g. Table ES, Box ES, Finding etc")
                objCaptionsMgr.deleteCustomCaptions()
                objCaptionsMgr.installCustomCaptions()
                'MsgBox("resetCaptions")

        End Select
        '
        objGlobals.glb_get_wrdApp.ScreenUpdating = True
        '
    End Sub


    Public Sub grp_Style_grpES_Click(sender As Object, e As RibbonControlEventArgs) Handles grpStylesES_StyleSet.Click, grpStylesES_Heading5_ES.Click, grpStylesES_Heading4_ES.Click, grpStylesES_Heading3_ES.Click, grpStylesES_Heading2_ES.Click
        Dim objSectMgr As New cSectionMgr()
        Dim objStyles As New cStylesManager()
        Dim objWrkAround As New cWorkArounds()
        Dim sect As Word.Section
        Dim strStyleName As String
        Dim myDoc As Word.Document
        '
        sect = objSectMgr.objGlobals.glb_get_wrdApp.Selection.Sections.Item(1)
        myDoc = sect.Range.Document
        '
        objSectMgr.objGlobals.glb_screen_update(False)
        '
        strStyleName = "Body Text"
        '
        Try
            Select Case e.Control.Id
                Case "grpStylesES_StyleSet"
                    Call objStyles.styles_format_styleSetES()
                Case "grpStylesES_Heading1_ES"
                    strStyleName = "Heading 1 (ES)"
                    objStyles.applyStyleToSelection(strStyleName)
                Case "grpStylesES_Heading2_ES"
                    strStyleName = "Heading 2 (ES)"
                    objStyles.applyStyleToSelection(strStyleName)
                Case "grpStylesES_Heading3_ES"
                    strStyleName = "Heading 3 (ES)"
                    objStyles.applyStyleToSelection(strStyleName)
                Case "grpStylesES_Heading4_ES"
                    strStyleName = "Heading 4 (ES)"
                    objStyles.applyStyleToSelection(strStyleName)
                Case "grpStylesES_Heading5_ES"
                    strStyleName = "Heading 5 (ES)"
                    objStyles.applyStyleToSelection(strStyleName)
                Case "grpStylesES_Heading6_ES"
                    strStyleName = "Heading 6 (ES)"
                    objStyles.applyStyleToSelection(strStyleName)
            End Select
        Catch ex As Exception

        End Try
        '
        objSectMgr.objGlobals.glb_screen_update(True)
        '
        objWrkAround.wrk_fix_forCursorRace()
        'objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = True
        '
    End Sub
    '

    Private Sub PIF_Styl_grpStylesText_Click(sender As Object, e As RibbonControlEventArgs) Handles grpStylesText_BodyText.Click
        Dim objStyles As New cStylesManager()
        Dim objSectMgr As New cSectionMgr()
        Dim objTools As New cTools()
        Dim strStyleName As String
        '
        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = False
        strStyleName = "Normal"
        '
        Select Case e.Control.Id
            Case "grpStylesText_BodyText"
                Call objStyles.applyStyleToSelection("Body Text")
            Case "grpStylesText_PasteAsText"
                objTools.tools_paste_AsUnformattedText()
            Case "grpStylesText_FormatPainter"
                objSectMgr.objGlobals.glb_get_wrdApp.Dialogs(WdWordDialog.wdDialogFormatPicture).Show()
        End Select

        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = True
        '
    End Sub
    '
    Private Sub PIF_Styl_grpRpt_Click(sender As Object, e As RibbonControlEventArgs) Handles grpStylesRpt_Intro.Click
        Dim objStyles As New cStylesManager()
        Dim objSectMgr As New cSectionMgr()
        Dim objRptMgr As New cReport()
        Dim objLstStyles As New clstStyles()
        Dim objTOCMgr As New cTOCMgr()

        Dim strStyleName, strColourPickerMode As String
        Dim objWrkAround As New cWorkArounds()
        Dim myDoc As Word.Document
        '
        objSectMgr.objGlobals.glb_screen_update(False)
        '
        'objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = False
        myDoc = objSectMgr.objGlobals.glb_get_wrdActiveDoc
        strStyleName = "Normal"
        strColourPickerMode = ""
        '
        Select Case e.Control.Id
            Case "grpTest_test_listStyles_Heading_to_NoNum"
                'objStylesMgr.style_lstLevel_removeNumberingHeadings_Main(objGlobals.glb_get_wrdActiveDoc)
                objLstStyles.lstStyle_set_HeadingsBD_noNumbered(objSectMgr.objGlobals.glb_get_wrdActiveDoc)
                objLstStyles.lstStyle_set_HeadingsAP_noNumbered(objSectMgr.objGlobals.glb_get_wrdActiveDoc)
                objTOCMgr.toc_update_TOCs(objSectMgr.objGlobals.glb_get_wrdActiveDoc)
                'j = 1
            Case "grpTest_test_listStyles_Heading_to_Numbered"
                objLstStyles.lstStyle_set_HeadingsBD_Numbered(objSectMgr.objGlobals.glb_get_wrdActiveDoc)
                objLstStyles.lstStyle_set_HeadingsAP_Numbered(objSectMgr.objGlobals.glb_get_wrdActiveDoc)
                objTOCMgr.toc_update_TOCs(objSectMgr.objGlobals.glb_get_wrdActiveDoc)
                'j = 1

            Case "grpTblsPlh_adjustStyles"
                'objStyles.style_upgrade_Styles(myDoc, "dummy")
                'MsgBox("here")
                objRptMgr.Rpt_styles_Upgrade_for_ReportType(myDoc, "dummy")
                objSectMgr.objGlobals.glb_screen_update(True)
                '
            Case "grpStylesRpt_get_pictControl", "grpCpImages_Custom_backcolour"
                Select Case e.Control.Id
                    Case "grpStylesRpt_get_pictControl"
                        strColourPickerMode = "text_Colour"
                    Case "grpCpImages_Custom_backcolour"
                        strColourPickerMode = "backPanel"
                End Select
                objSectMgr.objGlobals.glb_show_ColorPicker(strColourPickerMode)
                'frm = New frm_colorPicker(strColourPickerMode)
                'frm.Show()
                '
            Case "grpStylesRpt_col_ApplyWhite"
                objStyles.applyColourToSelection("white")
            Case "grpStylesRpt_col_ApplyGrey"
                objStyles.applyColourToSelection("grey")
            Case "grpStylesRpt_col_ApplySecondaryPurple"
                objStyles.applyColourToSelection("purple_Secondary")
            Case "grpStylesRpt_col_ApplyPurple"
                objStyles.applyColourToSelection("purple")
            Case "grpStylesRpt_col_ReApplyBase"
                objStyles.applyColourToSelection("reset")
        End Select
        '
        Select Case e.Control.Id
            Case "grpStylesRpt_Intro"
                strStyleName = "Introduction"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Glossary"
                strStyleName = "Glossary"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Emphasis"
                strStyleName = "SideNote (Regular Left)"
                objStyles.applyStyleToSelection(strStyleName)

        End Select

        Select Case e.Control.Id
            Case "grpStylesRpt_StyleSet"
                objStyles.styles_format_styleSetRpt()
            Case "grpStylesRpt_Heading1_Rpt"
                strStyleName = "Heading 1"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Heading2_Rpt"
                strStyleName = "Heading 2"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Heading3_Rpt"
                strStyleName = "Heading 3"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Heading4_Rpt"
                strStyleName = "Heading 4"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Heading5_Rpt"
                strStyleName = "Heading 5"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Heading6_Rpt"
                strStyleName = "Heading 6"
                objStyles.applyStyleToSelection(strStyleName)
        End Select
        '
        Select Case e.Control.Id
            Case "grpStylesRpt_HeadingNoNum_StyleSet"
                objStyles.styles_format_styleSet_NoNumber()
            Case "grpStylesRpt_Heading1NoNum_Rpt"
                strStyleName = "Heading 1 (no number)"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Heading2NoNum_Rpt"
                strStyleName = "Heading 2 (no number)"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Heading3NoNum_Rpt"
                strStyleName = "Heading 3 (no number)"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Heading4NoNum_Rpt"
                strStyleName = "Heading 4 (no number)"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Heading5NoNum_Rpt"
                strStyleName = "Heading 5 (no number)"
                objStyles.applyStyleToSelection(strStyleName)

        End Select
        '
        '*** Fix for cursor race case by Word interaction with fields in footer.
        '*** A Word issue not a me issue
        '
        objSectMgr.objGlobals.glb_screen_update(True)
        '
        objWrkAround.wrk_fix_forCursorRace()

        objSectMgr.objGlobals.glb_screen_update(True)
        '
        'objWrkAround.wrk_fix_forCursorRace()

        'objWrkAround.glb_screen_update(True)
        'objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = True
        '
    End Sub
    '
    Private Sub PIF_Styl_grpStylesLists_Click(sender As Object, e As RibbonControlEventArgs) Handles grpStylesLists_ListNumber3.Click, grpStylesLists_ListNumber2.Click, grpStylesLists_ListNumber1.Click, grpStylesLists_List3.Click, grpStylesLists_List2.Click, grpStylesLists_List1.Click
        Dim objStyles As New cStylesManager()
        Dim objSectMgr As New cSectionMgr()
        Dim strStyleName As String
        '
        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = False
        strStyleName = "Normal"
        '
        Select Case e.Control.Id
            Case "grpStylesLists_List1"
                objStyles.applyStyleToSelection("List Bullet")
            Case "grpStylesLists_List2"
                objStyles.applyStyleToSelection("List Bullet 2")
            Case "grpStylesLists_List3"
                objStyles.applyStyleToSelection("List Bullet 3")
            Case "grpStylesLists_ListNumber1"
                objStyles.applyStyleToSelection("List Number")
            Case "grpStylesLists_ListNumber2"
                objStyles.applyStyleToSelection("List Number 2")
            Case "grpStylesLists_ListNumber3"
                objStyles.applyStyleToSelection("List Number 3")

        End Select

        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = True

    End Sub

    Private Sub PIF_Styl_grpStylesOther_Click(sender As Object, e As RibbonControlEventArgs) Handles grpStylesOther_QuoteSource.Click, grpStylesOther_QuoteBlt.Click, grpStylesOther_Quote.Click
        Dim objStyles As New cStylesManager()
        Dim objSectMgr As New cSectionMgr()
        Dim strStyleName As String
        '
        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = False
        strStyleName = "Normal"
        '
        Select Case e.Control.Id
            Case "grpStylesOther_Quote"
                objStyles.applyStyleToSelection("Quote")
            Case "grpStylesOther_QuoteBlt"
                objStyles.applyStyleToSelection("Quote List Bullet")
            Case "grpStylesOther_QuoteSource"
                objStyles.applyStyleToSelection("Quote Source")
        End Select
        '
        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = True
        '
    End Sub
    '

    Private Sub PIF_Styl_grpStylesPullOuts_Click(sender As Object, e As RibbonControlEventArgs) Handles grpPullouts_emphasisBox_Right.Click, grpPullouts_emphasisBox_Left.Click, grpPullouts_emphasisBox_Centre.Click, grpPullouts_emphasisBox_TextStyle_Right_2.Click, grpPullouts_emphasisBox_TextStyle_Left_2.Click, grpPullouts_emphasisBox_TextStyle_Centre_2.Click
        Dim objStyles As New cStylesManager()
        Dim objSectMgr As New cSectionMgr()
        Dim objGlobals As New cGlobals()
        Dim strStyleName, strMsg As String
        '
        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = False
        strStyleName = "Normal"
        strMsg = ""
        '
        Select Case e.Control.Id
            Case "grpStylesPullouts_text", "grpPullouts_text", "grpPullouts_emphasisBox_Left", "grpPullouts_emphasisBox_Left_2"
                'This is the the standard, italic with right justified text
                strMsg = objStyles.styl_insert_EmphasisBox("left")
                If strMsg <> "" Then
                    MsgBox(strMsg)
                End If
                'shp = objGlobals.glb_get_wrdDoc.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, -122.0, 20.0, 102.95, 42.8, objGlobals.glb_get_wrdSelRng())
                'shp.Fill.ForeColor.RGB = RGB(255, 0, 0)
                'rng.Paragraphs.ad
                '
            Case "grpPullouts_emphasisBox_Centre", "grpPullouts_emphasisBox_Centre_2"
                strMsg = objStyles.styl_insert_EmphasisBox("centre")
                If strMsg <> "" Then
                    MsgBox(strMsg)
                End If
                '
            Case "grpPullouts_emphasisBox_Right", "grpPullouts_emphasisBox_Right_2"
                strMsg = objStyles.styl_insert_EmphasisBox("right")
                If strMsg <> "" Then
                    MsgBox(strMsg)
                End If

            Case "grpPullouts_emphasisBox_TextStyle_Left", "grpPullouts_emphasisBox_TextStyle_Left_2"
                objStyles.applyStyleToSelection("Emphasis Text (Left)")

            Case "grpPullouts_emphasisBox_TextStyle_Centre", "grpPullouts_emphasisBox_TextStyle_Centre_2"
                objStyles.applyStyleToSelection("Emphasis Text (Centre)")

            Case "grpPullouts_emphasisBox_TextStyle_Right", "grpPullouts_emphasisBox_TextStyle_Right_2"
                objStyles.applyStyleToSelection("Emphasis Text (Right)")

                '
            Case "grpPullouts_text_jr_italic"
                objStyles.applyStyleToSelection("SideNote (Italic Right)")
            Case "grpStylesPullouts_text_jl_italic"
                objStyles.applyStyleToSelection("SideNote (Italic Left)")
            Case "grpStylesPullouts_text_jr_Not_italic"
                objStyles.applyStyleToSelection("SideNote (Regular Right)")
            Case "grpStylesPullouts_text_jl_Not_italic"
                objStyles.applyStyleToSelection("SideNote (Regular Left)")
            Case "grpStylesPullouts_text_jl_Not_italic_dash"
                objStyles.applyStyleToSelection("SideNote Dash")
            Case "grpStylesPullouts_text_jl_Not_italic_bullet"
                objStyles.applyStyleToSelection("SideNote Bullet")
                '
            Case "grpStylesPullouts_pict", "grpPullouts_pict"
                strMsg = objStyles.styl_insert_EmphasisBox("pict")
                If strMsg <> "" Then
                    MsgBox(strMsg)
                End If
        End Select

        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = True

    End Sub
    '

    Private Sub grp_Style_grpRpt_Click(sender As Object, e As RibbonControlEventArgs) Handles grpStylesRpt_StyleSet.Click, grpStylesRpt_HeadingNoNum_StyleSet.Click, grpStylesRpt_Heading5NoNum_Rpt.Click, grpStylesRpt_Heading5_Rpt.Click, grpStylesRpt_Heading4NoNum_Rpt.Click, grpStylesRpt_Heading4_Rpt.Click, grpStylesRpt_Heading3NoNum_Rpt.Click, grpStylesRpt_Heading3_Rpt.Click, grpStylesRpt_Heading2NoNum_Rpt.Click, grpStylesRpt_Heading2_Rpt.Click, grpStyles_mnu_Heading3Numbering_btn_on.Click, grpStyles_mnu_Heading3Numbering_btn_off.Click
        Dim objStyles As New cStylesManager()
        Dim objSectMgr As New cSectionMgr()
        Dim objRptMgr As New cReport()
        Dim objLstStyles As New clstStyles()
        Dim objTOCMgr As New cTOCMgr()

        Dim strStyleName, strColourPickerMode As String
        Dim objWrkAround As New cWorkArounds()
        'Dim frm As frm_colorPicker
        Dim myDoc As Word.Document
        '
        objSectMgr.objGlobals.glb_screen_update(False)
        '
        'objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = False
        myDoc = objSectMgr.objGlobals.glb_get_wrdActiveDoc
        strStyleName = "Normal"
        strColourPickerMode = ""
        '
        Select Case e.Control.Id
            Case "grpTest_test_listStyles_Heading_to_NoNum", "grpStyles_mnu_Heading3Numbering_btn_off"
                'objStylesMgr.style_lstLevel_removeNumberingHeadings_Main(objGlobals.glb_get_wrdActiveDoc)
                objLstStyles.lstStyle_set_HeadingsBD_noNumbered(objSectMgr.objGlobals.glb_get_wrdActiveDoc)
                objLstStyles.lstStyle_set_HeadingsAP_noNumbered(objSectMgr.objGlobals.glb_get_wrdActiveDoc)
                objTOCMgr.toc_update_TOCs(objSectMgr.objGlobals.glb_get_wrdActiveDoc)
                'j = 1
            Case "grpTest_test_listStyles_Heading_to_Numbered", "grpStyles_mnu_Heading3Numbering_btn_on"
                objLstStyles.lstStyle_set_HeadingsBD_Numbered(objSectMgr.objGlobals.glb_get_wrdActiveDoc)
                objLstStyles.lstStyle_set_HeadingsAP_Numbered(objSectMgr.objGlobals.glb_get_wrdActiveDoc)
                objTOCMgr.toc_update_TOCs(objSectMgr.objGlobals.glb_get_wrdActiveDoc)
                'j = 1

            Case "grpTblsPlh_adjustStyles"
                'objStyles.style_upgrade_Styles(myDoc, "dummy")
                'MsgBox("here")
                objRptMgr.Rpt_styles_Upgrade_for_ReportType(myDoc, "dummy")
                objSectMgr.objGlobals.glb_screen_update(True)
                '
            Case "grpStylesRpt_get_pictControl", "grpCpImages_Custom_backcolour"
                Select Case e.Control.Id
                    Case "grpStylesRpt_get_pictControl"
                        strColourPickerMode = "text_Colour"
                    Case "grpCpImages_Custom_backcolour"
                        strColourPickerMode = "backPanel"
                End Select
                objSectMgr.objGlobals.glb_show_ColorPicker(strColourPickerMode)
                'frm = New frm_colorPicker(strColourPickerMode)
                'frm.Show()
                '
            Case "grpStylesRpt_col_ApplyWhite"
                objStyles.applyColourToSelection("white")
            Case "grpStylesRpt_col_ApplyGrey"
                objStyles.applyColourToSelection("grey")
            Case "grpStylesRpt_col_ApplySecondaryPurple"
                objStyles.applyColourToSelection("purple_Secondary")
            Case "grpStylesRpt_col_ApplyPurple"
                objStyles.applyColourToSelection("purple")
            Case "grpStylesRpt_col_ReApplyBase"
                objStyles.applyColourToSelection("reset")
        End Select
        '
        Select Case e.Control.Id
            Case "grpStylesRpt_Intro"
                strStyleName = "Introduction"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Glossary"
                strStyleName = "Glossary"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Emphasis"
                strStyleName = "SideNote (Regular Left)"
                objStyles.applyStyleToSelection(strStyleName)

        End Select

        Select Case e.Control.Id
            Case "grpStylesRpt_StyleSet"
                objStyles.styles_format_styleSetRpt()
            Case "grpStylesRpt_Heading1_Rpt"
                strStyleName = "Heading 1"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Heading2_Rpt"
                strStyleName = "Heading 2"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Heading3_Rpt"
                strStyleName = "Heading 3"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Heading4_Rpt"
                strStyleName = "Heading 4"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Heading5_Rpt"
                strStyleName = "Heading 5"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Heading6_Rpt"
                strStyleName = "Heading 6"
                objStyles.applyStyleToSelection(strStyleName)
        End Select
        '
        Select Case e.Control.Id
            Case "grpStylesRpt_HeadingNoNum_StyleSet"
                objStyles.styles_format_styleSet_NoNumber()
            Case "grpStylesRpt_Heading1NoNum_Rpt"
                strStyleName = "Heading 1 (no number)"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Heading2NoNum_Rpt"
                strStyleName = "Heading 2 (no number)"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Heading3NoNum_Rpt"
                strStyleName = "Heading 3 (no number)"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Heading4NoNum_Rpt"
                strStyleName = "Heading 4 (no number)"
                objStyles.applyStyleToSelection(strStyleName)
            Case "grpStylesRpt_Heading5NoNum_Rpt"
                strStyleName = "Heading 5 (no number)"
                objStyles.applyStyleToSelection(strStyleName)

        End Select
        '
        '*** Fix for cursor race case by Word interaction with fields in footer.
        '*** A Word issue not a me issue
        '
        objSectMgr.objGlobals.glb_screen_update(True)
        '
        objWrkAround.wrk_fix_forCursorRace()

        objSectMgr.objGlobals.glb_screen_update(True)
        '
        'objWrkAround.wrk_fix_forCursorRace()
    End Sub
    '
    Private Sub PIF_PgS_fixesGroup_Click(sender As Object, e As RibbonControlEventArgs) Handles grpFixes_Repairs_delSpace1_betweenWords.Click, grpFixes_Repairs_delSpace1_atSentenceEnd.Click, grpFixes_ScreenUpdatingOn.Click, grpFixes_ScreenUpdatingOff.Click, grpFixes_Repairs_SetLanguage.Click, grpFixes_Repairs_remSpaces_indrCells.Click, grpFixes_Repairs_remCharChar.Click, grpFixes_RePaginate.Click, grpFixes_PaginateOff.Click
        Dim objRptMgr As New cReport()
        Dim objSectMgr As New cSectionMgr()
        Dim objColsMgr As New cColsHandler()
        Dim objToolsMgr As New cTools()
        Dim strRptMode As String
        Dim sect As Word.Section
        '
        sect = objSectMgr.objGlobals.glb_get_wrdApp.Selection.Sections.Item(1)
        'objSectMgr.objGlobals.glb_screen_update(False)
        '
        strRptMode = objRptMgr.Rpt_Mode_Get()

        Select Case e.Control.Id
            Case "grpFixes_ScreenUpdatingOff"
                MsgBox("Screen updating is about to be turned off")
                objSectMgr.objGlobals.glb_screen_update(False)
            Case "grpFixes_ScreenUpdatingOn"
                MsgBox("Screen updating is on")
                objSectMgr.objGlobals.glb_screen_update(True)
            Case "grpFixes_Repairs_delSpace1_betweenWords"
                MsgBox("This method scans the entire document and ensures that there is only one space between words")
                objToolsMgr.spaces_OneBetweenWords()
            Case "grpFixes_Repairs_delSpace1_atSentenceEnd"
                MsgBox("This method scans the entire document and ensures that there is only one space after each sentence")
                objToolsMgr.tools_Remove_TrailingSpacesFromParagraphs(objSectMgr.objGlobals.glb_get_wrdActiveDoc)
                'objToolsMgr.spaces_One()
            Case "xgrpFixes_Repairs_delSpace1_atSentenceEnd"
                MsgBox("This method scans the entire document and ensures that there is only one space after each sentence")
                objToolsMgr.DeleteExtraSpacesInTable()
            Case "grpFixes_Repairs_SetLanguage"                     'changes language
                MsgBox("This method will turn 'proofing on' and set the language to English")
                objToolsMgr.SpellCheckProofing(True)

        End Select
        '
        'objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = True

    End Sub

    Private Sub PIF_PgS_grpSectOptions_Click(sender As Object, e As RibbonControlEventArgs) Handles grpRpt_sectOptions_btn_delSection.Click, grpSectOptions_sect_InsertSectionBounded_Prt_wide.Click, grpSectOptions_sect_InsertSectionBounded_Prt.Click, grpSectOptions_sect_InsertSection_AtSelection.Click, grpSectOptions_sect_InsertSectionBounded_Lnd_wide.Click, grpSectOptions_sect_InsertSectionBounded_Lnd.Click, grpSectOptions_header_ClearTextandShapes.Click, grpSectOptions_footer_ClearText.Click, grpSectOptions_footer_ClearTextandPageNum.Click, grpSectOptions_footer_clearSubTitleField.Click, grpSectOptions_footer_resetText.Click, grpSectOptions_hfs_restoreHF_RP.Click, grpSectOptions_hfs_restoreHF_ES.Click, grpSectOptions_hfs_restoreHF_AP.Click, grpSectOptions_resetTo_Prt_RP.Click, grpSectOptions_resetTo_Prt_ES.Click, grpSectOptions_resetTo_Prt_AP.Click, grpSectOptions_resetTo_Lnd_RP.Click, grpSectOptions_resetTo_Lnd_ES.Click, grpSectOptions_resetTo_Lnd_AP.Click, grpSectOptions_resizeTo_Landscape.Click, grpSectOptions_resize_toggleWidth.Click, grpSectOptions_resizeTo_Portrait.Click, grpSectOptions_sect_InsertSection_InFront.Click, grpSectOptions_sect_InsertSection_Behind.Click
        Dim objGlobals As New cGlobals()
        Dim objChptMgr As New cChptBase()
        Dim objTagsMgr As New cTagsMgr()
        Dim objSectMgr As cSectionMgr
        Dim objIsOKMgr As cIsOKToDo
        Dim objHfMgr As cHeaderFooterMgr
        Dim objMsgMgr As New cMessageManager()
        Dim obkWrkAround As New cWorkArounds()
        Dim objPgNumMgr As New cPageNumberMgr()
        Dim objViewMgr As cViewManager
        Dim objRptMgr As New cReport()
        Dim tokens() As String
        'Dim objCpMgr As New cCoverPageMgr()
        Dim marginOffset As Single
        Dim rng As Word.Range
        Dim sect As Word.Section
        Dim strMsg, strHeaderStyleName, strPgNumFormat As String
        '
        'sect = Globals.ThisDocument.Application.Selection.Sections.Item(1)
        strMsg = ""
        strHeaderStyleName = "spacer"
        sect = objGlobals.glb_get_wrdApp.Selection.Sections.Item(1)
        objGlobals.glb_screen_update(False)
        '
        Try
            Select Case e.Control.Id
                Case "grpSectOptions_ToggleWidth"
                    objSectMgr = New cSectionMgr()
                    marginOffset = objGlobals.glb_get_TableOutdent()
                    objSectMgr.sect_Toggle_Width(sect, marginOffset)
                Case "grpSectOptions_header_ClearTextandShapes"
                    objHfMgr = New cHeaderFooterMgr()
                    objHfMgr.hf_headers_DeleteContents_All(sect)
                Case "grpSectOptions_footer_ClearText"
                    objHfMgr = New cHeaderFooterMgr()
                    objHfMgr.hf_footers_DeleteContents_Text(sect)

                Case "grpSectOptions_footer_clearSubTitleField"
                    objIsOKMgr = New cIsOKToDo()
                    objHfMgr = New cHeaderFooterMgr()
                    '
                    For Each sect In objGlobals.glb_get_wrdActiveDoc().Sections
                        If objIsOKMgr.isOKto_reset_footerText(sect) = "" Then
                            'objHfMgr.hf_footers_DeleteContents_Text(sect)
                            objHfMgr.hf_footer_deleteEmptySubTitleField(sect)

                            'objHfMgr.hf_insert_coverPageStyleRefs(tbl, doMirror)
                        End If
                    Next
                    '
                    MsgBox("The Sub Title field has been removed from all footers")
                    '

                Case "grpSectOptions_footer_resetText"
                    objIsOKMgr = New cIsOKToDo()
                    objHfMgr = New cHeaderFooterMgr()
                    '
                    For Each sect In objGlobals.glb_get_wrdActiveDoc().Sections
                        If objIsOKMgr.isOKto_reset_footerText(sect) = "" Then
                            'objHfMgr.hf_footers_DeleteContents_Text(sect)
                            objHfMgr.hf_footer_addFooterFields(sect)

                            'objHfMgr.hf_insert_coverPageStyleRefs(tbl, doMirror)
                        End If
                    Next
                    '
                    MsgBox("The footer 'text' reset is complete")
                    '
                    '
                Case "grpSectOptions_footer_ClearTextandPageNum"
                    objHfMgr = New cHeaderFooterMgr()
                    objHfMgr.hf_footers_DeleteContents_All(sect)
                Case "grpSectOptions_hfs_deleteHeader"
                    objHfMgr = New cHeaderFooterMgr()
                    objHfMgr.hf_headers_delete(sect)
                Case "grpSectOptions_hfs_deleteFooter"
                    objHfMgr = New cHeaderFooterMgr()
                    objHfMgr.hf_footers_delete(sect)
                Case "grpSectOptions_hfs_deleteAll"
                    objHfMgr = New cHeaderFooterMgr()
                    objHfMgr.hf_hfs_deleteAll(sect)
                Case "grpSectOptions_hfs_restoreHF_ES", "grpSectOptions_hfs_restoreHF_RP", "grpSectOptions_hfs_restoreHF_AP"
                    objHfMgr = New cHeaderFooterMgr()
                    objIsOKMgr = New cIsOKToDo()
                    '
                    tokens = Split(e.Control.Id, "_")
                    Select Case tokens.Last
                        Case "ES"
                            strHeaderStyleName = objTagsMgr.bnr_get_tagStyles(objTagsMgr.tag_es)
                        Case "RP"
                            strHeaderStyleName = objTagsMgr.bnr_get_tagStyles(objTagsMgr.tag_chpt_body)
                        Case "AP"
                            strHeaderStyleName = objTagsMgr.bnr_get_tagStyles(objTagsMgr.tag_chpt_AP)
                    End Select
                    '
                    If objIsOKMgr.isOKto_doAction_inReportBody() = objIsOKMgr._isOK Then
                        objHfMgr.hf_headers_insert(objGlobals.glb_get_wrdSect(),,,,, strHeaderStyleName)
                        objHfMgr.hf_footers_insert(objGlobals.glb_get_wrdSect())
                        MsgBox("The section headers and footers have been restored")
                    Else
                        MsgBox(objIsOKMgr.isOKto_doAction_inReportBody())
                    End If

                    '
                    objGlobals.glb_screen_update(True)
                    '
                Case "grpTest_hfs_insertHeaders"
                    objHfMgr = New cHeaderFooterMgr()
                    objHfMgr.hf_headers_insert(sect)
                Case "grpTest_hfs_insertFooters"
                    objHfMgr = New cHeaderFooterMgr()
                    objHfMgr.hf_footers_delete(sect)

                Case "grpTest_hfs_linkToPrevious_No"
                    objHfMgr = New cHeaderFooterMgr()
                    objHfMgr.hf_hfs_linkUnlinkAll(sect, False)
                Case "grpTest_hfs_linkToprevious_Yes"
                    objHfMgr = New cHeaderFooterMgr()
                    objHfMgr.hf_hfs_linkUnlinkAll(sect, True)
            End Select
            '
            Select Case e.Control.Id
                Case "grpSectOptions_resizeTo_Portrait"
                    objSectMgr = New cSectionMgr()
                    objViewMgr = New cViewManager()
                    '
                    'objViewMgr.vw_change_ColumnsAndRows(sect, 1, 1)
                    objViewMgr.vw_fit_fullPage(sect)
                    objSectMgr.sct_resize_ToPortrait(sect)
                    'objSectMgr.sct_reset_ToPortrait(sect)
                    '
                    objGlobals.glb_screen_update(True)
                    '
                Case "grpSectOptions_resizeTo_Portrait_wide"
                    objSectMgr = New cSectionMgr()
                    objViewMgr = New cViewManager()
                    '
                    'objViewMgr.vw_change_ColumnsAndRows(sect, 1, 1)
                    objViewMgr.vw_fit_fullPage(sect)
                    objSectMgr.sct_resize_ToPortrait(sect)
                    '
                    'marginOffset = objGlobals.glb_get_TableOutdent()
                    'objSectMgr.sect_Toggle_Width(sect, marginOffset)
                    '
                    objGlobals.glb_screen_update(True)

                Case "grpSectOptions_resizeTo_Landscape"
                    objSectMgr = New cSectionMgr()
                    objViewMgr = New cViewManager()
                    '
                    objViewMgr.vw_fit_fullPage(sect)
                    objSectMgr.sct_resize_ToLandscape(sect)
                    'marginOffset = objGlobals.glb_get_TableOutdent()
                    'objSectMgr.sect_Toggle_Width(sect, marginOffset)

                    '
                    objGlobals.glb_screen_update(True)
                    '
                Case "grpSectOptions_resizeTo_Landscape_wide"
                    objSectMgr = New cSectionMgr()
                    objViewMgr = New cViewManager()
                    '
                    objViewMgr.vw_fit_fullPage(sect)
                    objSectMgr.sct_resize_ToLandscape(sect)
                    '
                    'marginOffset = objGlobals.glb_get_TableOutdent()
                    'objSectMgr.sect_Toggle_Width(sect, marginOffset)
                    '
                    objGlobals.glb_screen_update(True)
                    '
                Case "grpSectOptions_resize_toggleWidth"
                    objSectMgr = New cSectionMgr()
                    objViewMgr = New cViewManager()
                    '
                    objViewMgr.vw_fit_fullPage(sect)
                    'objSectMgr.sct_resize_ToLandscape(sect)
                    '
                    marginOffset = objGlobals.glb_get_TableOutdent()
                    objSectMgr.sect_Toggle_Width(sect, marginOffset)
                    '
                    objGlobals.glb_screen_update(True)
                    '

                    '
                Case "grpSectOptions_resetTo_Prt_ES", "grpSectOptions_resetTo_Prt_RP", "grpSectOptions_resetTo_Prt_AP"
                    objSectMgr = New cSectionMgr()
                    objHfMgr = New cHeaderFooterMgr()
                    objIsOKMgr = New cIsOKToDo()
                    '
                    tokens = Split(e.Control.Id, "_")
                    Select Case tokens.Last
                        Case "ES"
                            objHfMgr.hf_hfs_linkUnlinkAll(sect, False)
                            strHeaderStyleName = objTagsMgr.bnr_get_tagStyles(objTagsMgr.tag_es)
                            objChptMgr.chptBase_PageNumbering_Set(sect, False, 1, "es")

                        Case "RP"
                            objHfMgr.hf_hfs_linkUnlinkAll(sect, False)
                            strHeaderStyleName = objTagsMgr.bnr_get_tagStyles(objTagsMgr.tag_chpt_body)
                            strPgNumFormat = objPgNumMgr.pgNum_get_numFormat_ForDoc()
                            objChptMgr.chptBase_PageNumbering_Set(sect, False, 1, strPgNumFormat)

                        Case "AP"
                            objHfMgr.hf_hfs_linkUnlinkAll(sect, False)
                            strHeaderStyleName = objTagsMgr.bnr_get_tagStyles(objTagsMgr.tag_chpt_AP)
                            objChptMgr.chptBase_PageNumbering_Set(sect, True, 1, "ap")

                    End Select
                    '
                    If objIsOKMgr.isOKto_doAction_inReportBody() = objIsOKMgr._isOK Then
                        objSectMgr.sct_reset_ToPortrait(sect, strHeaderStyleName)
                        MsgBox("Reset to Portrait is complete")
                    Else
                        MsgBox(objIsOKMgr.isOKto_doAction_inReportBody())
                    End If

                Case "grpSectOptions_resetTo_Lnd_ES", "grpSectOptions_resetTo_Lnd_RP", "grpSectOptions_resetTo_Lnd_AP"
                    objSectMgr = New cSectionMgr()
                    objHfMgr = New cHeaderFooterMgr()
                    objIsOKMgr = New cIsOKToDo()
                    '
                    tokens = Split(e.Control.Id, "_")
                    Select Case tokens.Last
                        Case "ES"
                            objHfMgr.hf_hfs_linkUnlinkAll(sect, False)
                            strHeaderStyleName = objTagsMgr.bnr_get_tagStyles(objTagsMgr.tag_es)
                            objChptMgr.chptBase_PageNumbering_Set(sect, False, 1, "es")

                        Case "RP"
                            objHfMgr.hf_hfs_linkUnlinkAll(sect, False)
                            strHeaderStyleName = objTagsMgr.bnr_get_tagStyles(objTagsMgr.tag_chpt_body)
                            strPgNumFormat = objPgNumMgr.pgNum_get_numFormat_ForDoc()
                            objChptMgr.chptBase_PageNumbering_Set(sect, False, 1, strPgNumFormat)

                        Case "AP"
                            objHfMgr.hf_hfs_linkUnlinkAll(sect, False)
                            strHeaderStyleName = objTagsMgr.bnr_get_tagStyles(objTagsMgr.tag_chpt_AP)
                            objChptMgr.chptBase_PageNumbering_Set(sect, True, 1, "ap")

                    End Select
                    '
                    If objIsOKMgr.isOKto_doAction_inReportBody() = objIsOKMgr._isOK Then
                        objSectMgr.sct_reset_ToLandscape(sect,, strHeaderStyleName)
                        MsgBox("Reset to Landscape is complete")
                    Else
                        MsgBox(objIsOKMgr.isOKto_doAction_inReportBody())
                    End If

            End Select
            '

            '
            Select Case e.Control.Id
                Case "grpSectOptions_sect_InsertSectionBounded_Lnd", "grpSectOptions_sect_InsertSectionBounded_Lnd_wide", "grpExecSum_Landscape", "grpReport_Landscape_body"
                    'Inserts a bounded section at the current selection. When finished the selection is at the 
                    'beginning of the nhew bounded section
                    objSectMgr = New cSectionMgr()
                    If Not objSectMgr.sct_Sel_IsIn_Or_JustUnderTable() Then
                        rng = objGlobals.glb_get_wrdSelRng()
                        sect = objSectMgr.sct_insert_SectionBounded(rng, "std_Lnd", 6, "newPage", False)
                        'sect = rng.Document.Sections.Item(sect.Index + 1)
                        '
                        'For brief we need to ensure that all sections except the first and last are set to 
                        'diff first page = false
                        'objSectMgr.sct_setSection_toNoDiffFirstPage(rng.Document)
                        objSectMgr.sct_adjustInsertedSections_forBrief(sect)
                        '
                        Select Case e.Control.Id
                            Case "grpSectOptions_sect_InsertSectionBounded_Lnd_wide"
                                'If the 'wide', 'full width' option has been chosen then we need to toggle
                                'the width of the new landscape section to 'full width'
                                marginOffset = objGlobals.glb_get_TableOutdent()
                                'objSectMgr.sect_Toggle_Width(rng.Document.Sections.Item(sect.Index - 1), marginOffset)
                                objSectMgr.sect_Toggle_Width(sect, marginOffset)

                        End Select
                        '
                        'sect = rng.Document.Sections.Item(sect.Index - 1)
                        rng = sect.Range
                        rng.Collapse(WdCollapseDirection.wdCollapseStart)
                        rng.Select()
                    Else
                        objMsgMgr.msg_insertionPoint_IsIn_Or_JustUnderATable()
                    End If
                    '
                    obkWrkAround.wrk_fix_forCursorRace()
                    '
                Case "grpSectOptions_sect_InsertSectionBounded_Prt", "grpSectOptions_sect_InsertSectionBounded_Prt_wide"
                    Dim objInserTestMgr As New cInsertTestMgr()
                    objSectMgr = New cSectionMgr()
                    'Inserts a bounded section at the current selection. When finished the selection is at the 
                    'beginning of the nhew bounded section
                    '
                    'If isLandscape And Section Is banner etc then don't do it 
                    strMsg = objInserTestMgr.ins_is_OKToInsert(objGlobals.glb_get_wrdSect, True)
                    If strMsg = "" Then
                        If Not objSectMgr.sct_Sel_IsIn_Or_JustUnderTable() Then
                            rng = objGlobals.glb_get_wrdSelRng()
                            sect = objSectMgr.sct_insert_SectionBounded(rng, "std_Prt", 6, "newPage", False)
                            '
                            'For brief we need to ensure that all sections except the first and last are set to 
                            'diff first page = false
                            'objSectMgr.sct_setSection_toNoDiffFirstPage(rng.Document)
                            '
                            '****
                            objSectMgr.sct_adjustInsertedSections_forBrief(sect)
                            '
                            '**
                            '
                            Select Case e.Control.Id
                                Case "grpSectOptions_sect_InsertSectionBounded_Prt_wide"
                                    marginOffset = objGlobals.glb_get_TableOutdent()
                                    objSectMgr.sect_Toggle_Width(sect, marginOffset)
                            End Select
                            '
                            'sect = rng.Document.Sections.Item(sect.Index - 1)
                            rng = sect.Range
                            rng.Collapse(WdCollapseDirection.wdCollapseStart)
                            rng.Select()
                        Else
                            objMsgMgr.msg_insertionPoint_IsIn_Or_JustUnderATable()
                        End If
                    Else
                        MsgBox(strMsg)
                    End If
                '
                Case "grpSectOptions_sect_InsertSection_AtSelection"
                    'Goes to the beginning of the current section and inserts at that point. The selection remains
                    'at the same point and in the same section it was in
                    objSectMgr = New cSectionMgr()
                    'MsgBox("Section")'
                    If Not objSectMgr.sct_Sel_IsIn_Or_JustUnderTable() Then
                        sect = objSectMgr.sct_insert_SectionAtSelection()
                        Select Case objRptMgr.Rpt_Mode_Get()
                            Case objRptMgr.rpt_isBrief
                                sect.PageSetup.DifferentFirstPageHeaderFooter = False
                            Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd

                        End Select
                        '
                        objGlobals.glb_screen_update(True)
                        '
                    Else
                        objMsgMgr.msg_insertionPoint_IsIn_Or_JustUnderATable()
                    End If
                    '
                    objGlobals.glb_screen_update(True)
                    '
                Case "grpSectOptions_sect_InsertSection_InFront"
                    'Goes to the beginning of the current section and inserts at that point
                    objSectMgr = New cSectionMgr()
                    sect = objGlobals.glb_get_wrdSect()
                    sect = objSectMgr.sct_insert_Section(False, sect, 6, "newPage", False, "flow")
                    'sect = objSectMgr.sct_insert_SectionInFront(objGlobals.glb_get_wrdApp.Selection.Range, 6, "newPage", False)
                    'sect = objSectMgr.sct_insert_SectionAtSelection(objGlobals.glb_get_wrdApp.Selection.Range,, "oddPage")
                    rng = sect.Range
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    rng.Select()

                Case "grpSectOptions_sect_InsertSection_Behind"
                    'Goes to the beginning of the current section and inserts at that point
                    objSectMgr = New cSectionMgr()
                    sect = objGlobals.glb_get_wrdSect()
                    'rng = objGlobals.glb_get_wrdApp.Selection.Range
                    sect = objSectMgr.sct_insert_Section(True, sect, 6, "newPage", False, "flow")

                    'sect = objSectMgr.sct_insert_SectionBehind(objGlobals.glb_get_wrdApp.Selection.Range, 6, "newPage", False)
                    '
                    'rng = sect.Range
                    'rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    'rng.Select()
                '
                Case "grpSectOptions_sect_DeleteSection", "grpRpt_sectOptions_btn_delSection"
                    objSectMgr = New cSectionMgr()
                    objSectMgr.sct_delete_Section(sect)
            End Select
            '
        Catch ex As Exception
            objGlobals.glb_screen_update(True)
        End Try
        '
        'objGlobals.glb_get_wrdApp.ScreenUpdating = True
        objGlobals.glb_screen_update(True)
        '
        obkWrkAround.wrk_fix_forCursorRace()

    End Sub

    Private Sub PIF_Pgs_grpGlossaryEtc_Click(sender As Object, e As RibbonControlEventArgs) Handles grpRpt_btn_GlossaryAndAbbreviations_bblk.Click, grpRpt_btn_GlossaryAndAbbreviations.Click, grpOther_worksCited_bblk.Click, grpOther_worksCited.Click, grpOther_references_bblk.Click, grpOther_references.Click, grpOther_bibliography_bblk.Click, grpOther_bibliography.Click
        Dim objGlobals As New cGlobals()
        Dim objGlosMgr As New cGlossary()
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objIsOK As New cIsOKToDo()
        Dim objRptMgr As New cReport()
        Dim objBBlkMgr As New cBBlocksHandler()
        Dim objChptBase As New cChptBase()
        Dim objBnrMgr As New cChptBanner()
        Dim lstOfCpgCnFTOC As New Collection()
        Dim sectIndex As Integer
        Dim sect As Word.Section
        Dim myDoc As Word.Document
        Dim fld As Word.Field
        Dim rng As Word.Range
        Dim strRptMode, strTagStyle As String
        Dim placeBehind As Boolean
        '
        objGlobals.glb_get_wrdApp.ScreenUpdating = False
        myDoc = objGlobals.glb_get_wrdActiveDoc
        strRptMode = objRptMgr.Rpt_Mode_Get()
        strTagStyle = ""
        '
        placeBehind = True
        sect = objGlobals.glb_get_wrdSect()
        '
        Select Case e.Control.Id
            Case "grpRpt_btn_GlossaryAndAbbreviations_bblkx"

            Case "grpRpt_btn_GlossaryAndAbbreviations", "grpRpt_btn_GlossaryAndAbbreviations_bblk"
                Select Case strRptMode
                    Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                        lstOfCpgCnFTOC = objHfMgr.hf_getTagStyleMap_CpContactsFrontAndTOC(myDoc)
                        'objRptMgr.chptBase_PageNumbering_Set(sect, True, 1, "es")
                        '
                        If lstOfCpgCnFTOC.Count > 0 Then
                            Try
                                'Get the section.index of the TOC
                                sectIndex = CInt(lstOfCpgCnFTOC.Item("toc"))
                                Select Case e.Control.Id
                                    Case "grpRpt_btn_GlossaryAndAbbreviations"
                                        'Software build option
                                        objGlosMgr.glos_insert_Glossary(Not placeBehind, myDoc.Sections.Item(sectIndex + 1))
                                    Case "grpRpt_btn_GlossaryAndAbbreviations_bblk"
                                        'Building Blocks option.. faster
                                        sect = myDoc.Sections.Item(sectIndex + 1)
                                        rng = sect.Range
                                        rng.Collapse(WdCollapseDirection.wdCollapseStart)
                                        rng.Select()
                                        rng = objBBlkMgr.bblk_insert_sectionInFront("aa_Glossary")
                                        objGlosMgr.glos_select_GlossaryFirstEntry(rng.Sections.First)

                                End Select

                            Catch ex As Exception
                                Try
                                    'Look for the front contacts page.. If found got to the next section and insert before
                                    sectIndex = CInt(lstOfCpgCnFTOC.Item("cnf"))
                                    sect = myDoc.Sections.Item(sectIndex + 1)
                                    '
                                    Select Case e.Control.Id
                                        Case "grpRpt_btn_GlossaryAndAbbreviations"
                                            objGlosMgr.glos_insert_Glossary(Not placeBehind, sect)
                                        Case "grpRpt_btn_GlossaryAndAbbreviations_bblk"
                                            'Building Blocks option.. faster
                                            rng = sect.Range
                                            rng.Collapse(WdCollapseDirection.wdCollapseStart)
                                            rng.Select()
                                            rng = objBBlkMgr.bblk_insert_sectionInFront("aa_Glossary")
                                            objGlosMgr.glos_select_GlossaryFirstEntry(rng.Sections.First)

                                    End Select

                                Catch ex1 As Exception
                                    Try
                                        'Let's try putting it behind the cover page
                                        sectIndex = CInt(lstOfCpgCnFTOC.Item("cpg"))
                                        sect = myDoc.Sections.Item(sectIndex + 1)

                                        Select Case e.Control.Id
                                            Case "grpRpt_btn_GlossaryAndAbbreviations"
                                                objGlosMgr.glos_insert_Glossary(Not placeBehind, sect)
                                            Case "grpRpt_btn_GlossaryAndAbbreviations_bblk"
                                                'Building Blocks option.. faster
                                                rng = sect.Range
                                                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                                                rng.Select()
                                                rng = objBBlkMgr.bblk_insert_sectionInFront("aa_Glossary")
                                                objGlosMgr.glos_select_GlossaryFirstEntry(rng.Sections.First)
                                        End Select
                                    Catch ex2 As Exception

                                    End Try

                                End Try

                            End Try
                        Else
                            MsgBox("To insert a 'Glossary', the document needs a cover page, a front contacts page or a table of contents")
                        End If
                        '
                    Case objRptMgr.rpt_isBrief
                        If objGlobals.glb_selection_IsInTable() Then
                            MsgBox("Your cursor needs to be clear of any tables")
                        Else
                            objGlosMgr.glos_insert_Glossary(Not placeBehind, sect)
                        End If
                End Select
                '
            Case "grpOther_bibliography", "grpOther_bibliography_bblk"
                Select Case e.Control.Id
                    Case "grpOther_bibliography"
                        'Software build
                        objGlosMgr.glos_insert_Biblio(Not placeBehind, objGlobals.glb_get_wrdSect)
                    Case "grpOther_bibliography_bblk"
                        'Building Block build
                        rng = sect.Range
                        rng.Collapse(WdCollapseDirection.wdCollapseStart)
                        rng.Select()
                        rng = objBBlkMgr.bblk_insert_sectionInFront("aa_Biblio")


                End Select
            Case "grpOther_references", "grpOther_references_bblk"
                Select Case e.Control.Id
                    Case "grpOther_references"
                        objGlosMgr.glos_insert_Refs(Not placeBehind, objGlobals.glb_get_wrdSect)
                    Case "grpOther_references_bblk"
                        'Building Block build
                        rng = sect.Range
                        rng.Collapse(WdCollapseDirection.wdCollapseStart)
                        rng.Select()
                        rng = objBBlkMgr.bblk_insert_sectionInFront("aa_Biblio")
                        '
                        strTagStyle = objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_glos_refsCited)
                        objHfMgr.hf_tags_setTagStyle(rng.Sections.First, strTagStyle)

                        rng = objChptBase.chptBase_getRange_Heading1(rng.Sections.First)
                        rng.Text = "References"
                        rng.Select()
                        '
                        sect = rng.Sections.First
                        rng = sect.Range
                        '
                        If rng.Sections.First.Range.Fields.Count <> 0 Then
                            fld = rng.Sections.First.Range.Fields.Item(1)
                            fld.Select()
                        End If
                        '
                End Select
            Case "grpOther_worksCited", "grpOther_worksCited_bblk"
                Select Case e.Control.Id
                    Case "grpOther_worksCited"
                        'Software build
                        objGlosMgr.glos_insert_WorksCited(Not placeBehind, objGlobals.glb_get_wrdSect)
                    Case "grpOther_worksCited_bblk"
                        'Building Block build
                        rng = sect.Range
                        rng.Collapse(WdCollapseDirection.wdCollapseStart)
                        rng.Select()
                        rng = objBBlkMgr.bblk_insert_sectionInFront("aa_Biblio")
                        '
                        strTagStyle = objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_glos_wrks)
                        objHfMgr.hf_tags_setTagStyle(rng.Sections.First, strTagStyle)
                        '
                        rng = objChptBase.chptBase_getRange_Heading1(rng.Sections.First)
                        rng.Text = "Works Cited"
                        rng.Select()
                        '
                        sect = rng.Sections.First
                        rng = sect.Range
                        '
                        If rng.Sections.First.Range.Fields.Count <> 0 Then
                            fld = rng.Sections.First.Range.Fields.Item(1)
                            fld.Select()
                        End If
                        '
                End Select


        End Select
        '
        objGlobals.glb_get_wrdApp.ScreenUpdating = True

    End Sub

    Private Sub PIF_PgS_grpImageHandling_Click(sender As Object, e As RibbonControlEventArgs) Handles grpReport_mnu01_ImageBackPanel.Click, grpImageHandling_insert_BackPanel.Click, grpImageHandling_delete_BackPanel.Click, submnu_SetTransparency_to_75.Click, submnu_SetTransparency_to_50.Click, submnu_SetTransparency_to_25.Click, submnu_SetTransparency_to_100.Click, submnu_SetTransparency_to_0.Click, mnu_SetBackPanel_to_BannerHeight.Click, grpImageHandling_Reset_backcolour_to_CaseStudyGrey.Click, grpImageHandling_Reset_backcolour.Click, grpImageHandling_BackPanelFill_RawImageFromFile.Click, grpImageHandling_BackPanelFill_FromFile.Click, grpImageHandling_BackPanelFill_FromClipBoard.Click, grpImageHandling_Custom_backcolour.Click, grpWCAG_mnu_SetTransparency_to_75.Click, grpWCAG_mnu_SetTransparency_to_50.Click, grpWCAG_mnu_SetTransparency_to_25.Click, grpWCAG_mnu_SetTransparency_to_100.Click, grpWCAG_mnu_SetTransparency_to_0.Click
        Dim objImgGetEdit As New cImageGetAndEdit()
        Dim lstOfBackPanels As List(Of cShapeMgr)
        Dim objGlobals As New cGlobals()
        Dim objkWrkAround As New cWorkArounds()
        Dim objBackPanelMgr As New cBackPanelMgr()
        Dim objBrfMgr As New cReportBrief()
        Dim objShpMgr As New cShapeMgr()
        'Dim frmPicker As frm_colorPicker
        Dim tokens() As String
        Dim strResult As String
        Dim transparency As Single

        Dim objisOKToDo As New cIsOKToDo()
        Dim shp As Word.Shape
        Dim hf As HeaderFooter
        Dim sect As Word.Section
        Dim strMsg, strErrorMsg, strColourPickerMode As String
        '
        sect = objGlobals.glb_get_wrdSect()
        strErrorMsg = "This section does not have a standard image back panel." + vbCrLf + vbCrLf + "You'll need to add one using the 'Insert image back panel' menu item in the menu group below"

        '
        'These routines will handle panels etc in the Primary unless there is a first page
        'In which case, the first page becomes the default
        '
        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        If sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Exists Then
            hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        End If
        '
        '
        '*** Problem with the location of panel test
        Select Case e.Control.Id
            Case "grpImageHandling_BackPanelFill_FromFile"
                'The imgGet method also contain the error messages
                objImgGetEdit.imgGet_fill_backPanelFromFile_aac_BackColour()
                '
            Case "grpImageHandling_BackPanelFill_FromClipBoard"
                '
                objImgGetEdit.imgGet_fill_backPanelFromClipboard_aac_BackColour()
                '
            Case "grpImageHandling_BackPanelFill_RawImageFromFile"
                '
                objImgGetEdit.imgGet_fill_withRawImageFromFile_aac_BackColour()
                '
            Case "grpImageHandling_Reset_backcolour"
                'objBackPanelMgr.pnl_reset_BackPanelColour(sect)
                objBackPanelMgr.pnl_reset_BackPanelColour(sect)

                'lstOfBackPanels = objBackPanelMgr.pnl_getBackPanel_PlaceHolders(sect)                     'To get rid of any existing back panels
                '
                'If lstOfBackPanels.Count <> 0 Then
                'objShpMgr = lstOfBackPanels.Item(0)
                'objBackPanelMgr.pnl_reset_BackPanelColour(objShpMgr)
                'Else
                'MsgBox(strErrorMsg)
                'End If
                '
                objkWrkAround.wrk_fix_forCursorRace()
                '
            Case "grpImageHandling_Reset_backcolour_to_CaseStudyGrey"
                objBackPanelMgr.pnl_reset_BackPanelColour(sect,, objGlobals._glb_colour_CaseStudy_Grey)

            Case "grpImageHandling_Custom_backcolour", "grpReport_mnu01_ImageBackPanel"
                strColourPickerMode = "backPanel"
                '
                If objBackPanelMgr.pnl_has_BackPanel(sect) Then
                    objGlobals.glb_show_ColorPicker(strColourPickerMode)
                    'frmPicker = New frm_colorPicker(strColourPickerMode)
                    'frmPicker.Show()
                Else
                    MsgBox(strErrorMsg)
                End If
                objkWrkAround.wrk_fix_forCursorRace()
                '
        End Select

        '
        Select Case e.Control.Id
            Case "grpImageHandling_insert_BackPanel"
                If Not objBackPanelMgr.pnl_has_BackPanel(sect) Then
                    shp = objBackPanelMgr.pnl_BackPanel_Insert(hf, RGB(210, 191, 229))
                Else
                    MsgBox("This section already had a back panel",, "Back Panel Warning")
                End If
                '
            Case "grpImageHandling_delete_BackPanel"
                '
                strMsg = objisOKToDo.isOKto_Insert_BackPanel(sect)
                '
                If strMsg = "" Then
                    If objBackPanelMgr.pnl_has_BackPanel(sect) Then
                        'Dim lstOfPanels As New List(Of cShapeMgr)
                        '
                        'lstOfPanels = objBackPanelMgr.pnl_getBackPanel_PlaceHolders(sect)
                        'objBackPanelMgr.pnl_BackPanel_Delete(lstOfPanels)
                        objBackPanelMgr.pnl_BackPanel_Delete(sect)
                        '
                    Else
                        MsgBox("This section does not have an image back panel.")
                    End If
                Else
                    MsgBox(strMsg,, "Back Panel Warning")
                End If


                '
        End Select
        '
        Select Case e.Control.Id
            Case "submnu_SetTransparency_to_0", "submnu_SetTransparency_to_25", "submnu_SetTransparency_to_50", "submnu_SetTransparency_to_75", "submnu_SetTransparency_to_100"
                tokens = Split(e.Control.Id, "_")
                strResult = tokens.Last
                '
                transparency = CSng(strResult) / 100.0
                '
                If objBackPanelMgr.pnl_has_BackPanel(sect) Then
                    objBackPanelMgr.pnl_reset_BackPanelTransparency(transparency, sect)
                    MsgBox("Transparency settings complete")
                Else
                    MsgBox("This section does not have an image back panel.")
                End If
                '

            Case "grpWCAG_mnu_SetTransparency_to_0", "grpWCAG_mnu_SetTransparency_to_25", "grpWCAG_mnu_SetTransparency_to_50", "grpWCAG_mnu_SetTransparency_to_75", "grpWCAG_mnu_SetTransparency_to_100"
                tokens = Split(e.Control.Id, "_")
                strResult = tokens.Last
                '
                transparency = CSng(strResult) / 100.0
                '
                MsgBox("This method will search the document for image back panels that may contain images or solid colour. Any that are found wii have their transparency set to a value of your choice.")
                '
                For Each sect In objGlobals.glb_get_wrdActiveDoc.Sections
                    If objBackPanelMgr.pnl_has_BackPanel(sect) Then
                        objBackPanelMgr.pnl_reset_BackPanelTransparency(transparency, sect)
                    End If
                Next sect
                '
                MsgBox("Transparency settings complete")
                '
            Case "mnu_SetBackPanel_to_BannerHeight"
                'sect = objGlobals.glb_get_wrdActiveDoc.Sections.First
                sect = objGlobals.glb_get_wrdSect()
                If sect.Index = 1 And objBrfMgr.brf_is_brief(sect.Range.Document) Then
                    lstOfBackPanels = objBackPanelMgr.pnl_getBackPanel_PlaceHolders(sect)
                    For Each objShpMgr In lstOfBackPanels
                        objShpMgr.shp.Height = sect.PageSetup.PageHeight * 0.2
                        'shp.ZOrder(MsoZOrderCmd.msoBringInFrontOfText)
                        objShpMgr.shp.ZOrder(MsoZOrderCmd.msoSendBehindText)
                        objShpMgr.shp.ZOrder(MsoZOrderCmd.msoSendToBack)

                    Next
                    MsgBox("The 'Brief' first page banner panel has been resized to it's default dimensions")
                Else
                    MsgBox("This function can only be used if your cursor is in the first page of an 'AA Brief'. It allows you to undo the accidental/unwanted insertion of a full page back panel.")
                End If
                'hf = objBackPanelMgr.pnl_BackPanelBriefFirstPage_Insert(sect)
                '
                objkWrkAround.wrk_fix_forCursorRace()

        End Select
        '
        objGlobals.glb_cursors_setToNormal()
        objGlobals.glb_screen_update(True)
        objkWrkAround.wrk_fix_forCursorRace()
        '
    End Sub

    Private Sub PIF_PgS_grpViewTools_Click(sender As Object, e As RibbonControlEventArgs) Handles grpViewTools_Refresh_Stationery_Ref.Click, grpViewTools_Refresh_mnu_TOC.Click, grpViewTools_Refresh_mnu_Parts.Click, grpViewTools_Refresh_mnu_Chapters.Click, grpViewTools_Refresh_mnu_Tables.Click, grpViewTools_Refresh_mnu_Figures.Click, grpViewTools_Refresh_mnu_Boxes.Click, grpViewTools_Refresh_mnu_All.Click, grpViewTools_Refresh_mnu_Every.Click, grpToc_TOC_update.Click, grp_PgNumMgmnt_ReNum_std.Click, grp_PgNumMgmnt_ReNum_2Part.Click, grp_Finalise_DoAll.Click, grp_Finalise_CrossRefError.Click, grp_Finalise_upDateCopyrightNotice.Click, grp_Finalise_updateFields.Click, grp_Finalise_setFootersToBold.Click, grp_Finalise_RefreshTOC.Click
        Dim objGlobals As New cGlobals()
        Dim objFlds As New cFieldsMgr()
        Dim objMsgMgr As New cMessageManager()
        Dim objTOCMgr As New cTOCMgr()
        Dim objSectMgr As New cSectionMgr()
        Dim objLegals As New cLegalAndAbout()
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objCaptionsMgr As New cCaptionManager()
        Dim objTools As New cTools()
        Dim objProps As New cPropertyMgr()
        Dim objPgNumMgr As New cPageNumberMgr()
        Dim objRptMgr As New cReport()
        Dim rng As Word.Range
        Dim sect As Word.Section
        Dim myDoc As Word.Document
        Dim rslt As Boolean
        Dim frmErrors As frm_ListOfCrossRefErrors
        Dim strRptMode As String
        '
        objGlobals.glb_get_wrdApp.ScreenUpdating = False
        '
        sect = objGlobals.glb_get_wrdSect()
        myDoc = sect.Range.Document
        strRptMode = objRptMgr.Rpt_Mode_Get()
        '
        'MsgBox(control.Id)
        '
        Select Case e.Control.Id
            Case "grpViewTools_Refresh_Stationery_Ref"
                objFlds.flds_update_StyleRefs_Hfs()
                objMsgMgr.UpdateFunctionIsFinished()
                '
            Case "grpViewTools_Refresh_mnu_TOC", "grp_Finalise_RefreshTOC", "grpToc_TOC_update"
                objTOCMgr.toc_update_TOCs(myDoc)
                objTOCMgr.toc_upDate_TOFs()
                '
                objTOCMgr.toc_Selection_MoveToTOC(myDoc)
                '
                If e.Control.Id = "grpViewTools_Refresh_mnu_TOC" Then
                    objMsgMgr.UpdateFunctionIsFinished()
                End If
                '
            Case "grpViewTools_Refresh_mnu_Chapters"
                objFlds.updateSequenceNumbers_Chapters()
                objFlds.updateSequenceNumbers_Appendix()
                objFlds.updateSequenceNumbers_Tables()
                '
                objMsgMgr.UpdateFunctionIsFinished()
                '
            Case "grpViewTools_Refresh_mnu_Parts"
                objFlds.updateSequenceNumbers_Parts()
                objMsgMgr.UpdateFunctionIsFinished()
                '
            Case "grpViewTools_Refresh_mnu_Tables"
                objFlds.updateSequenceNumbers_Tables()
                objFlds.updateSequenceNumbers_Tables_AP()
                objFlds.updateSequenceNumbers_Tables_ES()
                objFlds.updateSequenceNumbers_Tables_LT()
                objMsgMgr.UpdateFunctionIsFinished()
                '
            Case "grpViewTools_Refresh_mnu_Figures"
                'objFlds.updateSequenceNumbers_Figures_WorkAround()
                objFlds.updateSequenceNumbers_Figures()
                objFlds.updateSequenceNumbers_Figures_Ap()
                objFlds.updateSequenceNumbers_Figures_ES()
                objMsgMgr.UpdateFunctionIsFinished()
                '
            Case "grpViewTools_Refresh_mnu_Boxes"
                objFlds.updateSequenceNumbers_Boxes()
                objFlds.updateSequenceNumbers_Boxes_Ap()
                objFlds.updateSequenceNumbers_Boxes_ES()
                objFlds.updateSequenceNumbers_Boxes_KeyFindings()
                objFlds.updateSequenceNumbers_Boxes_KeyFindings_ES()
                objFlds.updateSequenceNumbers_Boxes_Recommendation()
                objFlds.updateSequenceNumbers_Boxes_Recommendation_ES()
                objFlds.updateSequenceNumbers_Boxes_LT()
                objMsgMgr.UpdateFunctionIsFinished()
                '
            Case "grpViewTools_Refresh_mnu_All"
                objFlds.flds_update_StyleRefs_Hfs()
                objFlds.updateSequenceNumbers_All()
                objFlds.updateStyleRefs_All()
                '
                objTOCMgr.toc_update_TOCs(myDoc)
                objTOCMgr.toc_upDate_TOFs()
                objMsgMgr.UpdateFunctionIsFinished()
                '
            Case "grpViewTools_Refresh_mnu_Every", "grp_Finalise_updateFields"
                objFlds.flds_updateFields_All()
                '
            Case "grpViewTools_del_All"
                rslt = objMsgMgr.deleteAllMessage
                If rslt Then
                    objSectMgr.sct_delete_allSections()
                    objSectMgr.sct_reset_ToPortrait(objSectMgr.objGlobals.glb_get_wrdSect)
                    objMsgMgr.UpdateFunctionIsFinished()
                    '
                Else
                    'MsgBox ("No was chosen")
                End If

            Case "grpViewTools_smallBanner"
                'Not used 20210729
                '
                'If Globals.ThisDocument.Application.Selection.Range.Tables.Count >= 1 Then
                'tbl = Globals.ThisDocument.Application.Selection.Range.Tables(1)
                'Call objSectMgr.doBanner_Local(tbl)
                'Else

                'End If
            Case "grp_Finalise_CrossReferences"
                'Not used 20210729
                '
                'objFlds.CrossReference_changeStyle()
            Case "grp_Finalise_CrossRefError"
                'We now have a List of Fields that are in error
                'lst = objToolsMgr.CrossReference_ErrorsList()
                '
                frmErrors = New frm_ListOfCrossRefErrors(objGlobals.glb_get_wrdActiveDoc)
                frmErrors.TopMost = True
                frmErrors.Show()
                '
            Case "grp_Finalise_upDateCopyrightNotice"
                objLegals.legal_upDate_CopyRightNotice()
                objMsgMgr.UpdateFunctionIsFinished()
                '
            Case "grp_Finalise_setFootersToBold"
                objHfMgr.hf_footer_BoldStatus(objGlobals.glb_get_wrdActiveDoc, True)
                objMsgMgr.UpdateFunctionIsFinished()
                '
            Case "grp_Finalise_DoAll"
                Try
                    If objMsgMgr.doFinaliseMessage() Then
                        rng = objGlobals.glb_get_wrdSelRngAll()
                        '
                        objLegals.legal_upDate_CopyRightNotice()
                        '
                        'Now update the Fields
                        Call objFlds.flds_updateFields_All()
                        '
                        'Set footer to bold
                        objHfMgr.hf_footer_BoldStatus(objGlobals.glb_get_wrdActiveDoc, True)
                        'Set Cross Reference fields to bold or not bold
                        objCaptionsMgr.setFieldsBoldStatus(False)
                        'Now do one space only between words
                        Call objTools.spaces_OneBetweenWords()
                        'Now do one space after sentence end
                        Call objTools.tools_Remove_TrailingSpacesFromParagraphs(objGlobals.glb_get_wrdActiveDoc)
                        'Now finally do TOC
                        Call objTOCMgr.toc_update_TOCs(objGlobals.glb_get_wrdActiveDoc)
                        'Sticking, so do it twice as a work around
                        '*** AlexR - not necessary anymore So remove doubke updateTOC
                        'updateTOC now forces a refresh of the "App - Context" and "Chpt - Context" fields before updating the TOC
                        'Call objToolsMgr.updateTOCs
                        objTOCMgr.toc_upDate_TOFs()
                        '
                        MsgBox("Finalise Functions are complete")
                        '
                        rng.Select()
                    End If
                Catch ex As Exception

                End Try
                '
            Case "grp_PgNumMgmnt_ReNum_std"
                '
                objPgNumMgr.pgNum_set_numFormat_ForDoc("std")
                objPgNumMgr.pgNum_setBody_numFormat(objGlobals.glb_get_wrdActiveDoc())
                '
            Case "grp_PgNumMgmnt_ReNum_2Part"
                Select Case strRptMode
                    Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                        objPgNumMgr.pgNum_set_numFormat_ForDoc("2part")
                        objPgNumMgr.pgNum_setBody_numFormat(objGlobals.glb_get_wrdActiveDoc())
                        '
                    Case objRptMgr.rpt_isBrief
                        MsgBox("Two part page numbering is not supported in the AA Brief")
                End Select
        End Select
        '
        objGlobals.glb_get_wrdApp.ScreenUpdating = True
        '
    End Sub

    '
    Private Sub PIF_PgS_grpPageNumbering_Click(sender As Object, e As RibbonControlEventArgs) Handles grpFixes_RestartNumbering.Click, grpFixes_getNumberingDialog.Click, grpFixes_ContinueNumbering.Click, grpFixes_ApplyStdNumbering.Click, grpFixes_ApplyEsNumbering.Click, grpFixes_ApplyAppNumbering.Click
        Dim objSectMgr As New cSectionMgr()
        Dim objChpt As New cChptBase()
        Dim objMsgMgr As cMessageManager
        Dim objTOCMgr As New cTOCMgr()
        Dim pgNums As PageNumbers
        Dim sect As Word.Section
        Dim myDoc As Word.Document
        Dim dlg As Word.Dialog
        Dim dlgRslt As Integer
        '
        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = False
        '
        sect = objSectMgr.objGlobals.glb_get_wrdSelRng.Sections.Item(1)
        myDoc = sect.Range.Document

        pgNums = sect.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary).PageNumbers
        '
        'Page # Formatting options
        Select Case e.Control.Id
            Case "grpFixes_getNumberingDialog"
                'Need to set the Heading level to level 1 (i.e. 0) otherwise any other
                'changes cause a fault
                pgNums.HeadingLevelForChapter = 0
                dlg = Globals.ThisAddIn.Application.Dialogs.Item(WdWordDialog.wdDialogFormatPageNumber)
                dlgRslt = dlg.Show()
                'Dialogs(wdDialogFormatPageNumber).Show
            Case "grpFixes_ApplyStdNumbering"
                If Not objChpt.chptBase_PageNumbering_Set(sect, False, 1, "std") Then
                    objMsgMgr = New cMessageManager()
                    objMsgMgr.pageNumbers_Fault_ToESBodyChangeFailed("body")
                End If
                Call objTOCMgr.toc_update_TOCs(myDoc)
                '
            Case "grpFixes_ContinueNumbering"
                sect.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = False
                Call objTOCMgr.toc_update_TOCs(myDoc)
            'objSectMgr.currentSect.Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
            Case "grpFixes_RestartNumbering"
                sect.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
                sect.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
                Call objTOCMgr.toc_update_TOCs(myDoc)
            Case "grpFixes_ApplyEsNumbering"
                If Not objChpt.chptBase_PageNumbering_Set(sect, True, 1, "es") Then
                    objMsgMgr = New cMessageManager()
                    objMsgMgr.pageNumbers_Fault_ToESBodyChangeFailed("ES")
                End If
                Call objTOCMgr.toc_update_TOCs(myDoc)
                '
            Case "grpFixes_ApplyAppNumbering"
                If Not objChpt.chptBase_PageNumbering_Set(sect, True, 1, "ap") Then

                End If
                Call objTOCMgr.toc_update_TOCs(myDoc)

            Case Else
        End Select
        '
        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = True
        '
    End Sub

    Private Sub PIF_PgS_grpReport_Toggles_Click(sender As Object, e As RibbonControlEventArgs) Handles grpReport_btn_ToggleView.Click, tbHome_btn_ToggleView.Click
        Dim objSectMgr As New cSectionMgr()
        Dim objCpMgr As New cCoverPageMgr()
        Dim sect As Word.Section
        Dim objToolsMgr As New cTools()
        Dim myDoc As Word.Document
        '
        'objCpMgr = New cCoverPageMgr()
        'objSectMgr = New cSectionMgr()
        '
        myDoc = objSectMgr.objGlobals.glb_get_wrdActiveDoc()
        sect = objSectMgr.objGlobals.glb_get_wrdSect()

        Select Case e.Control.Id
            Case "grpReport_btn_ToggleView", "tbHome_btn_ToggleView"
                objToolsMgr.tools_viewHidden_Toggle()
                '
            Case "grpCpImages__ColourMode"
                If objCpMgr.cp_Bool_HasCoverPage(myDoc, sect) Then
                    'Call objCpMgr.cp_picture_changePictColour(sect, pressed)
                Else
                    MsgBox("Please make certian that you have a Cover Page With a picture")
                End If
            Case Else
        End Select
    End Sub
    '
    Private Sub PIF_PgS_grpWaterMark_Click(sender As Object, e As RibbonControlEventArgs) Handles grp_waterMark_removeStat_fromSect.Click, grp_waterMark_removeStat.Click, grp_waterMark_removeSec_fromSect.Click, grp_waterMark_removeSec.Click, grp_waterMark_removeAll.Click, grp_waterMark_draftOnly_add.Click, grp_waterMark_draft_add.Click, grp_waterMark_restricted_add.Click, grp_waterMark_confidential_add.Click, grp_waterMark_commercial_add.Click, grp_waterMark_cabinet_add.Click, grp_waterMark_atg_UNOFFICIAL_add.Click, grp_waterMark_atg_OFFICIAL_Sensitive_add.Click, grp_waterMark_atg_OFFICIAL_add.Click
        Dim objWMarksMgr As cWaterMarks
        Dim objWrkAround As New cWorkArounds()
        Dim objStylesMgr As New cStylesManager()
        Dim myDoc As Word.Document
        Dim rng As Word.Range
        Dim strMsg As String
        Dim sect As Section
        Dim objSectMgr As New cSectionMgr()
        Dim objGlobals As New cGlobals()
        '
        objWMarksMgr = New cWaterMarks
        '
        myDoc = objSectMgr.objGlobals.glb_get_wrdActiveDoc
        '
        If myDoc.ProtectionType <> WdProtectionType.wdNoProtection Then
            Call objWMarksMgr.msg_DocumentIsProtected()
            Exit Sub
        End If
        '
        objGlobals.glb_screen_update(False)
        '
        rng = objSectMgr.objGlobals.glb_get_wrdSelRng
        strMsg = "The selected water mark type will now be changed." + vbCrLf + vbCrLf _
            + "The time needed is dependent on the size of the document," + vbCrLf _
            + "the power of your PC, and the type of change requested." + vbCrLf + vbCrLf _
            + "Colour and alignment changes are the fastest (ocurring in seconds)." + vbCrLf _
            + "Adding or removing water marks may take anywhere from 30 seconds to a few minutes."
        '
        If e.Control.Id Like "grp_waterMark_*" Then
            MsgBox(strMsg)
        End If
        '
        objGlobals.glb_cursors_setToWait()
        '
        Try
            Select Case e.Control.Id
#Region "Release WaterMarks"
                Case "grp_waterMark_draft_add"
                    'This will remove legacy watermarks, as well as current
                    Call objWMarksMgr.waterMarks_Remove("*_stat")                      'Remove all Status Level Water Marks 
                    'Call objWMarksMgr.waterMarks_Remove("*_aa_stat")                    'Remove all Status Level Water Marks
                    Call objWMarksMgr.waterMarks_Add("draft_aa_stat")                      'Add the Draft Only Status Level Water Mark
                    strMsg = "The 'DRAFT' Water Mark" + vbCrLf _
                    + "has been successfully added to your document"
                Case "grp_waterMark_draftOnly_add"
                    'This will remove legacy watermarks, as well as current
                    Call objWMarksMgr.waterMarks_Remove("*_stat")                   'Remove all Status Level Water Marks
                    'Call objWMarksMgr.waterMarks_Remove("*_aa_stat")                   'Remove all Status Level Water Marks
                    Call objWMarksMgr.waterMarks_Add("draftOnly_aa_stat")                  'Add the Draft Only Status Level Water Mark
                    'Call objWMarksMgr.waterMarks_Add("draftOnlyBod_stat")              'Add the Draft Only Status Level Water Mark
                    strMsg = "The 'DRAFT ONLY' Water Mark" + vbCrLf _
                    + "has been successfully added to your document"
                Case "grp_waterMark_draftAlt1_add"
                    'This will remove legacy watermarks, as well as current
                    Call objWMarksMgr.waterMarks_Remove("*_stat")                   'Remove all Status Level Water Marks
                    'Call objWMarksMgr.waterMarks_Remove("*_aa_stat")                    'Remove all Status Level Water Marks
                    'Call objWMarksMgr.waterMarks_Add("draftReport_stat")               'Add the Draft Only Status Level Water Mark
                    Call objWMarksMgr.waterMarks_Add("draftOnlyBod_aa_stat")              'Add the Draft Only Status Level Water Mark
                    strMsg = "An 'Alternate Draft Only' " &
            "Water Mark" & vbCrLf & "has been successfully added to your document"
#End Region

#Region "Security Level Water Marks"

                Case "grp_waterMark_confidential_add"
                    'This will remove legacy watermarks, as well as current
                    Call objWMarksMgr.waterMarks_Remove("*_sec")                        'Remove all Security Level Water Marks
                    'Call objWMarksMgr.waterMarks_Remove("*_aa_sec")                     'Remove all Security Level Water Marks
                    Call objWMarksMgr.waterMarks_Add("Confidential_aa_sec")              'Add the Commercial-in-Confidence Security Level Water Mark
                    strMsg = "The 'Confidential' " &
            "Water Mark" & vbCrLf & "has been successfully added to your document"
                Case "grp_waterMark_commercial_add"
                    'This will remove legacy watermarks, as well as current
                    Call objWMarksMgr.waterMarks_Remove("*_sec")                        'Remove all Security Level Water Marks
                    'Call objWMarksMgr.waterMarks_Remove("*_aa_sec")                     'Remove all Security Level Water Marks
                    Call objWMarksMgr.waterMarks_Add("Commercial_aa_sec")              'Add the Commercial-in-Confidence Security Level Water Mark
                    strMsg = "The 'Commercial-in-Confidence' " &
            "Water Mark" & vbCrLf & "has been successfully added to your document"
                Case "grp_waterMark_cabinet_add"
                    'This will remove legacy watermarks, as well as current
                    Call objWMarksMgr.waterMarks_Remove("*_sec")                        'Remove all Security Level Water Marks
                    'Call objWMarksMgr.waterMarks_Remove("*_aa_sec")                     'Remove all Security Level Water Marks
                    Call objWMarksMgr.waterMarks_Add("Cabinet_aa_sec")
                    strMsg = "The 'Cabinet-in-Confidence' " &
            "Water Mark" & vbCrLf & "has been successfully added to your document"
                Case "grp_waterMark_restricted_add"
                    'This will remove legacy watermarks, as well as current
                    Call objWMarksMgr.waterMarks_Remove("*_sec")                        'Remove all Security Level Water Marks
                    'Call objWMarksMgr.waterMarks_Remove("*_aa_sec")                     'Remove all Security Level Water Marks
                    Call objWMarksMgr.waterMarks_Add("Restricted_aa_sec")
                    strMsg = "The 'Restricted Circulation' Water Mark" + vbCrLf _
                    + "has been successfully added to your document"
#End Region

#Region "Desemmination Limiting Markers - Attorney General Water Marks"
                Case "grp_waterMark_atg_UNOFFICIAL_add"
                    'This will remove legacy watermarks, as well as current
                    Call objWMarksMgr.waterMarks_Remove("*_sec")                        'Remove all Security Level Water Marks
                    'Call objWMarksMgr.waterMarks_Remove("*_aa_sec")                     'Remove all Security Level Water Marks

                    Call objWMarksMgr.waterMarks_Add("atg_UNOFFICIAL_aa_sec", "centre", RGB(255, 0, 0))            'Add the UNOFFICIAL Security Level Water Mark
                    strMsg = "The 'UNOFFICIAL' Water Mark" + vbCrLf _
                    + "has been successfully added to your document"

                Case "grp_waterMark_atg_OFFICIAL_add"
                    'This will remove legacy watermarks, as well as current
                    Call objWMarksMgr.waterMarks_Remove("*_sec")                                            'Remove all Security Level Water Marks
                    'Call objWMarksMgr.waterMarks_Remove("*_aa_sec")                                        'Remove all Security Level Water Marks
                    Call objWMarksMgr.waterMarks_Add("atg_OFFICIAL_aa_sec", "centre", RGB(255, 0, 0))              'Add the OFFICIAL Security Level Water Mark
                    strMsg = "The 'OFFICIAL' Water Mark" + vbCrLf _
                    + "has been successfully added to your document"

                Case "grp_waterMark_atg_OFFICIAL_Sensitive_add"
                    'This will remove legacy watermarks, as well as current
                    Call objWMarksMgr.waterMarks_Remove("*_sec")                        'Remove all Security Level Water Marks
                    'Call objWMarksMgr.waterMarks_Remove("*_aa_sec")                     'Remove all Security Level Water Marks

                    Call objWMarksMgr.waterMarks_Add("atg_OFFICIAL-Sensitive_aa_sec", "centre", RGB(255, 0, 0))    'Add the UNOFFICIAL:Sensitive Security Level Water Mark
                    strMsg = "The 'OFFICIAL:Sensitive' Water Mark" + vbCrLf _
                    + "has been successfully added to your document"

#End Region

#Region "Remove Water Marks"
                'Note that when a security level or status water mark is
                'removed, the associated style is refreshed. Tht is put back
                'to the way it was. In this way all insertions start off from
                'the same place
                Case "grp_waterMark_removeStat"
                    'Call objWMarksMgr.waterMarks_Remove("waterMark_aa_*_stat")                 'Remove All Status Water Marks
                    Call objWMarksMgr.waterMarks_Remove("*_stat")                               'Remove All Security Water 
                    'Call objWMarksMgr.waterMarks_Remove("*_aa_stat")                           'Remove All Security Water 
                    strMsg = "The Water Mark has been successfully removed"
                Case "grp_waterMark_removeSec"
                    'Call objWMarksMgr.waterMarks_Remove("waterMark_aa_*_sec")                  'Remove All Security Water Marks
                    Call objWMarksMgr.waterMarks_Remove("*_sec")                                'Remove All Security Water Marks
                    'Call objWMarksMgr.waterMarks_Remove("*_aa_sec")                             'Remove All Security Water Marks
                    strMsg = "The Water Mark has been successfully removed"
                Case "grp_waterMark_removeAll"
                    Call objWMarksMgr.waterMarks_Remove("waterMark_aa_*")                   'Remove ALL Water Marks
                    'Call objWMarksMgr.waterMarks_Remove("waterMark_aa_*_sec")                '
                    Call objWMarksMgr.waterMarks_Remove("*_aa_sec")                          'Remove All Security Water Marks
                    'Call objWMarksMgr.waterMarks_Remove("waterMark_aa_*_stat")
                    Call objWMarksMgr.waterMarks_Remove("*_aa_stat")                          'Remove All Document Status Water 
                    Call objWMarksMgr.waterMarks_Remove("*_AAC_version")               '
                    '
                    strMsg = "All Water Marks have been successfully removed"
            '
                Case "grp_waterMark_removeStat_fromSect"
                    objGlobals.glb_get_wrdSel()
                    sect = objGlobals.glb_get_wrdSel.Sections.Item(1)
                    Call objWMarksMgr.waterMarks_RemoveFromSection_Stat(sect)               'Remove Status Water Marks from current section
                    strMsg = "The Status Water Mark has been successfully removed"
                Case "grp_waterMark_removeSec_fromSect"
                    sect = objGlobals.glb_get_wrdSel.Sections.Item(1)
                    Call objWMarksMgr.waterMarks_RemoveFromSection_Sec(sect)                 'Remove Security Water Marks from current section
                    strMsg = "The Security Water Mark has been successfully removed"
                    '
#End Region

                Case Else

            End Select
            '
        Catch ex As Exception
            MsgBox("WaterMark error..Are you running the right template version? ")
        End Try
        '
        rng.Select()
        '
        objGlobals.glb_cursors_setToNormal()
        '
        objGlobals.glb_screen_update(True)
        '
        'objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = True
        'objSectMgr.objGlobals.glb_get_wrdApp.ScreenRefresh()
        '
        If strMsg <> "" Then MsgBox(strMsg)
        '
        objWrkAround.wrk_fix_forCursorRace()
        objGlobals.glb_screen_update(True)


    End Sub
    '
    Private Sub PIF_PgS_grpWaterMark_Simple_Click(sender As Object, e As RibbonControlEventArgs) Handles grp_waterMark_forceStat_StyleToDefault.Click, grp_waterMark_colour_red_stat.Click, grp_waterMark_colour_grey_stat.Click, grp_waterMark_NOTbold_sec.Click, grp_waterMark_colour_red_sec.Click, grp_waterMark_colour_grey_sec.Click, grp_waterMark_bold_sec.Click, grp_waterMark_alignment_Right_sec.Click, grp_waterMark_alignment_Centre_sec.Click, grp_waterMark_forceSec_StyleToDefault.Click
        Dim objWMarksMgr As cWaterMarks
        Dim objMsgMgr As New cMessageManager()
        Dim objStylesMgr As New cStylesManager()
        Dim dlgResult As Integer
        Dim rslt As Boolean
        Dim myDoc As Word.Document
        Dim rng As Word.Range
        Dim strMsg As String
        Dim objSectMgr As New cSectionMgr()
        Dim objGlobals As New cGlobals()
        '
        objWMarksMgr = New cWaterMarks
        rslt = False
        '
        myDoc = objSectMgr.objGlobals.glb_get_wrdActiveDoc
        '
        If myDoc.ProtectionType <> WdProtectionType.wdNoProtection Then
            Call objWMarksMgr.msg_DocumentIsProtected()
            Exit Sub
        End If
        '
        objGlobals.glb_screen_update(False)

        rng = objSectMgr.objGlobals.glb_get_wrdSelRng
        strMsg = ""
        '
        Select Case e.Control.Id
            Case "grp_waterMark_forceSec_StyleToDefault"
                strMsg = "This button will remove all security level water marks, before" + vbCrLf _
                    + "deleting and re-establishing the security level water mark style." + vbCrLf + vbCrLf _
                    + "Do you wish to continue?"
                dlgResult = MsgBox(strMsg, vbYesNo + vbDefaultButton2, "Water mark style reset")
                If dlgResult = vbYes Then
                    Try
                        objStylesMgr.style_waterMark_sec_ResetToDefault(myDoc)
                        MsgBox("The reset is complete.. You'll need to put your water marks back.")
                    Catch ex As Exception
                        strMsg = "There was an unknown problem with the reset." + vbCrLf _
                        + "You could try 'hand removing' the water mark style 'aa_waterMarkText_sec"
                        MsgBox(strMsg)
                    End Try
                End If

            '

            Case "grp_waterMark_forceStat_StyleToDefault"
                strMsg = "This button will remove all security level water marks, before" + vbCrLf _
                    + "deleting and re-establishing the security level water mark style." + vbCrLf + vbCrLf _
                    + "Do you wish to continue?"
                dlgResult = MsgBox(strMsg, vbYesNo + vbDefaultButton2, "Water mark style reset")
                If dlgResult = vbYes Then
                    Try
                        objStylesMgr.style_waterMark_stat_ResetToDefault(myDoc)
                        MsgBox("The reset is complete.. You'll need to put your water marks back.")
                    Catch ex As Exception
                        strMsg = "There was an unknown problem with the reset." + vbCrLf _
                        + "You could try 'hand removing' the water mark style 'aa_waterMarkText_stat"
                        MsgBox(strMsg)
                    End Try
                End If
            '
            Case "grp_waterMark_bold_sec"
                Try
                    myDoc.Styles.Item(objGlobals.glb_var_style_waterMark_sec).Font.Bold = True
                Catch ex As Exception
                    objMsgMgr.msgMgr_dlg_legacyWaterMarks()
                End Try
                '
                objGlobals.glb_screen_update(True)
                '
            Case "grp_waterMark_NOTbold_sec"
                Try
                    myDoc.Styles.Item(objGlobals.glb_var_style_waterMark_sec).Font.Bold = False
                Catch ex As Exception
                    objMsgMgr.msgMgr_dlg_legacyWaterMarks()
                End Try
                '
            Case "grp_waterMark_colour_red_sec"
                Try
                    myDoc.Styles.Item(objGlobals.glb_var_style_waterMark_sec).Font.Color = RGB(255, 0, 0)
                    MsgBox("Security status text in the header has been re-coloured to red")
                Catch ex As Exception
                    objMsgMgr.msgMgr_dlg_legacyWaterMarks()
                End Try
                '
            Case "grp_waterMark_colour_grey_sec"
                Try
                    myDoc.Styles.Item(objGlobals.glb_var_style_waterMark_sec).Font.Color = objGlobals._glb_colour_WaterMark_Grey_sec
                    MsgBox("Security status text in the header has been re-coloured to grey")
                    '
                Catch ex As Exception
                    objMsgMgr.msgMgr_dlg_legacyWaterMarks()
                End Try
            Case "grp_waterMark_colour_fromColorPicker_sec"

            Case "grp_waterMark_alignment_Centre_sec"
                rslt = objWMarksMgr.waterMark_sec_Alignment("centre", myDoc)
                If rslt Then
                    MsgBox("Security status text in the header has been centre aligned")
                Else
                    objMsgMgr.msgMgr_dlg_legacyWaterMarks()
                End If
                '
            Case "grp_waterMark_alignment_Right_sec"
                rslt = objWMarksMgr.waterMark_sec_Alignment("right", myDoc)
                '
                If rslt Then
                    MsgBox("Security status text in the header has been right aligned")
                Else
                    objMsgMgr.msgMgr_dlg_legacyWaterMarks()
                End If
                '
                '
            Case "grp_waterMark_bold_stat"
                Try
                    myDoc.Styles.Item(objGlobals.glb_var_style_waterMark_stat).Font.Bold = True
                Catch ex As Exception
                    objMsgMgr.msgMgr_dlg_legacyWaterMarks()
                End Try
                '
            Case "grp_waterMark_NOTbold_stat"
                Try
                    myDoc.Styles.Item(objGlobals.glb_var_style_waterMark_stat).Font.Bold = False
                Catch ex As Exception
                    objMsgMgr.msgMgr_dlg_legacyWaterMarks()
                End Try
                '
            Case "grp_waterMark_colour_red_stat"
                Try
                    myDoc.Styles.Item(objGlobals.glb_var_style_waterMark_stat).Font.Color = RGB(255, 0, 0)
                    MsgBox("Document status text ('DRAFT' etc) has been re-coloured to red")
                    '
                Catch ex As Exception
                    objMsgMgr.msgMgr_dlg_legacyWaterMarks()
                End Try
                '
            Case "grp_waterMark_colour_grey_stat"
                Try
                    myDoc.Styles.Item(objGlobals.glb_var_style_waterMark_stat).Font.Color = objGlobals._glb_colour_WaterMark_Grey_stat
                    MsgBox("Document status text ('DRAFT' etc) has been re-coloured to grey")

                Catch ex As Exception
                    objMsgMgr.msgMgr_dlg_legacyWaterMarks()
                End Try
                '
                '
            Case "grp_waterMark_colour_fromColorPicker_stat"

        End Select
        '
        objGlobals.glb_screen_update(True)
        '
    End Sub
    '
    Private Sub PIF_Pgs_grpWCAG_Click(sender As Object, e As RibbonControlEventArgs) Handles grpWCAG_convertThisDoc.Click, grpWCAG_notesOnAccessibility.Click
        Dim objFileMgr As New cFileHandler(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments))
        Dim objRptMgr As New cReport()
        Dim objPrint As cPrintAndDisplayServices
        Dim objGlobals As New cGlobals()
        Dim objTblsMgr As cTablesMgr
        Dim strResult, strDlgTitle, strItem, strPath As String
        'Dim doTablesAsOutdented, removeBanners, doAsNewFile As Boolean
        Dim listOfFiles As String()
        Dim singleList(1) As String
        'Dim dlgResult As DialogResult
        'Dim frmDocs As frm_SelectedDocs
        Dim frmPlacHolderList As frm_findTables
        Dim myDoc As Word.Document
        Dim objFormatMgr As New cFormatMgr()
        Dim objWCAGMgr As New cWCAGMgr()
        Dim mode As String
        Dim rng As Word.Range
        Dim objMsgMgr As New cMessageManager()
        Dim sel As Word.Selection
        Dim tbl As Word.Table
        Dim strMsg00, strMsg01, strMsg02 As String
        '
        strResult = ""
        strMsg02 = ""

        strDlgTitle = "Select the Word documents To be exported As partially WCAG compliant"
        strMsg00 = "Once you've converted all of the Floating PlaceHolders to inline," + vbCrLf + "try the conversion again to proceed to the next step." + vbCrLf + vbCrLf
        strMsg00 = strMsg00 + "Note, you can display the earlier dialog at any time by going to 'Placeholder > PlaceHolderMap."
        '
        strMsg01 = "Before conversion can begin the document must be checked for 'floating' PlaceHolders (since all PlaceHolders are made from tables" + vbCrLf
        strMsg01 = strMsg01 + "we are really looking for floating tables)." + vbCrLf + vbCrLf
        strMsg01 = strMsg01 + "Floating tables are not allowed in 'Accessible' documents." + vbCrLf + vbCrLf
        strMsg01 = strMsg01 + "If any are found you will be presented (after a few seconds) with" + vbCrLf
        strMsg01 = strMsg01 + "a tool that allows you to identify all floating tables in the " + vbCrLf
        strMsg01 = strMsg01 + "current document, and a means by which to adjust them to inline" + vbCrLf + vbCrLf
        strMsg01 = strMsg01 + "Inline PlaceHolders (tables) are OK for 'Accessibility" + vbCrLf + vbCrLf
        strMsg01 = strMsg01 + "Just be aware that adjusting a floating table to inline will change" + vbCrLf
        strMsg01 = strMsg01 + "the look of your document." + vbCrLf + vbCrLf
        strMsg01 = strMsg01 + "If you know that your document needs to be 'Accessible'. It's best" + vbCrLf
        strMsg01 = strMsg01 + "to avoid the use of floating PlaceHolders (tables) in the first place."
        '
        strMsg01 = "Before conversion the document will be checked for floating tables." + vbCrLf _
                    + "Floating tables are not allowed in Accessible documents and must" + vbCrLf _
                    + "be converted to inline tables." + vbCrLf + vbCrLf _
                    + "Since all PlaceHolders are made from tables, the check covers" + vbCrLf _
                    + "Figures, Boxes and Tables and can take anywhere between" + vbCrLf _
                    + "10 to 50 seconds depending on the size of the document." + vbCrLf + vbCrLf _
                    + "If any floating tables are found you will be automatically directed" + vbCrLf _
                    + "to a tool that lists all floating tables in the document and which" + vbCrLf _
                    + "will provide you with the means to adjust them to 'inline'." + vbCrLf + vbCrLf _
                    + "If none are found you'll be directed to the next step in the" + vbCrLf _
                    + "conversion process."
        '
        strMsg01 = "Please wait while the document is checked for floating tables." + vbCrLf _
                    + "Floating tables are not allowed in Accessible documents and must" + vbCrLf _
                    + "be converted to inline tables/placeholders." + vbCrLf + vbCrLf _
                    + "If any are found the 'Display Placeholders' tool will pop up" + vbCrLf _
                    + "(in 10 to 50 seconds). It will allow you to adjust these " + vbCrLf _
                    + "tables/placeholders to inline... Check for irregular tables" + vbCrLf _
        '
        strMsg02 = "This function will save and close your current document after creating copy in the same directory. The copy is the document that is converted. It's name is 'filename-wcag-yyyymmdd-hhmmss'." + vbCrLf + vbCrLf + "Since no two authors write a document in the same way it is not possible to convert a 'random' document to be fully 'Accessible'." + vbCrLf + vbCrLf + "The renamed document will be adjusted To maximise 'Accessibility Compliance'." + vbCrLf + vbCrLf + "The author will then need to use the various 'Accessibility Tools' in concert with Microsoft's 'Accessibility' tool at 'Review>Check Accessibility to identify and rectify those document elements that throw 'Accessibility' errors." + vbCrLf + vbCrLf + "If you know that your document will be sent to a client that requires it to be 'Accessible' then it's best to write it with that in mind." + vbCrLf + vbCrLf + "To that end Floating Tables and Tables with merged cells (referred to as irregular tables) will generate 'Accessibility' errors. The 'Placeholder Map' tool will let you easily find these in a large document."
        strItem = ""
        strPath = ""
        mode = ""
        '
        'Remember the click position that got us here
        Globals.ThisAddIn.point_PriorClick = System.Windows.Forms.Cursor.Position
        'Me.point_PriorClick = Cursor.Position
        Select Case e.Control.Id
            '
            'This is the only level of 'automatic' accessibility conversion that is offered
            'Accessiblity is too 'illdefined' and authors write documents in many different
            'ways that makes an automatic conversion just not possible... The solution
            'is education (i.e. write with accessibility in mind) and tools
            '
            Case "grpWCAG_convertThisDoc"
                'First check if there are floatig tables.. The author must deal with these
                '
                objTblsMgr = New cTablesMgr()
                '
                MsgBox(strMsg01,, "Notes on Floating tables in document")
                '
                If objTblsMgr.tbl_has_floatingTables(objGlobals.glb_get_wrdActiveDoc) Then
                    'Floating table warning
                    '
                    'Show placeholdere map and encourage the authr to deal with the floating tables
                    frmPlacHolderList = New frm_findTables()
                    frmPlacHolderList.ShowDialog()
                    MsgBox(strMsg00)
                    'GoTo finis
                Else
                    'MsgBox("Accessibility conversion Is currently being" + vbCrLf + "revised And will be available soon, 20240121" + vbCrLf + vbCrLf + "Try a document with floating tables (e.g. 'Emphasis')" + vbCrLf + "to see the 'floating table editor' in action")
                    'GoTo finis
                    '
                    objFileMgr.file_doc_toWCAG(objGlobals.glb_get_wrdActiveDoc)
                    '
                End If

            Case "grpWCAG_notesOnAccessibility"
                MsgBox(strMsg02, , "Notes on Accessibility")
                'Legacy hold over.. not used
            Case "grpWCAG_convertThisDoc_RemoveBanners"
                If objMsgMgr.msgMgr_dlg_doDocToWCAG() Then
                    objGlobals.glb_cursors_setToWait()
                    myDoc = objGlobals.glb_get_wrdActiveDoc
                    '
                    sel = objGlobals.glb_get_wrdSel
                    rng = sel.Range
                    tbl = objGlobals.glb_get_wrdSelTbl
                    If IsNothing(tbl) Then
                        objWCAGMgr.wcag_doc_ToWCAG(myDoc)
                        'Go back to where you started
                        rng.Select()
                        '
                    Else
                        objWCAGMgr.wcag_doc_ToWCAG(myDoc)
                        'Go to the first page
                        rng = myDoc.Sections.First.Range
                        rng.Collapse(WdCollapseDirection.wdCollapseStart)
                        rng.Select()
                        '
                    End If
                    '
                    objGlobals.glb_cursors_setToNormal()
                    objGlobals.glb_screen_update(True)
                    MsgBox("Conversion complete")
                    '
                    '
                Else

                End If
                '
                'Legacy hold over.. not used
            Case "grpWCAG_exportThisDoc"
                myDoc = objGlobals.glb_get_wrdActiveDoc
                '
                If objFileMgr.docSaveStatus(myDoc) Then
                    'Create an array of one document
                    listOfFiles = {myDoc.FullName}
                    '
                    'Save the exported version of myDoc to Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "\WCAG documents"
                    'frmDocs = New frm_SelectedDocs(myDoc, Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "\AAC WCAG conversion")
                    mode = "wcag"
                    'frmDocs = New frm_SelectedDocs(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "\AAC conversions - WCAG", mode)

                    'If frmDocs._createSaveDirectoriesIsOK Then
                    'frmDocs.setListOfDocs(listOfFiles)
                    'frmDocs.displayListOfDocs()
                    'frmDocs.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
                    'frmDocs.btn_Cancel.Select()
                    'frmDocs.Height = 360
                    '
                    'frmDocs.ShowDialog()
                    'Else
                    'MsgBox("Could Not create " + frmDocs._saveDirectoryFullName + " You may Not have the right permissions To Do this?")
                    'End If

                Else
                    MsgBox("The current document needs To be saved before it can be exported")
                End If
                '
                'Legacy hold over.. not used
            Case "grpWCAG_exportSelectedDocs"
                Try
                    listOfFiles = objFileMgr.getDocsToImport(strDlgTitle, , )
                    If Not IsNothing(listOfFiles) Then
                        mode = "wcag"
                        'frmDocs = New frm_SelectedDocs(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "\AAC conversions - WCAG", mode)
                        'If frmDocs._createSaveDirectoriesIsOK Then
                        'frmDocs.setListOfDocs(listOfFiles)
                        'frmDocs.displayListOfDocs()
                        'frmDocs.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
                        'frmDocs.Location = Me.point_PriorClick
                        'frmDocs.btn_Cancel.Select()
                        '
                        'frmDocs.ShowDialog()
                        'frmDocs.Show()
                        'Else
                        'MsgBox("Could Not create " + frmDocs._saveDirectoryFullName + " You may Not have the right permissions To Do this?")
                        'End If
                    Else
                        MsgBox("No documents selected")
                    End If

                Catch ex As Exception
                    MsgBox("Try/Catch fault In grpWCAG_exportSelectedDocs")
                End Try

            Case "grpWCAG_resetThisDoc", "grpStylesTools_resetStyle"
                rng = objGlobals.glb_get_wrdSel.Range
                Try
                    objRptMgr.Rpt_Styles_resetStyles_fromTemplate()
                    MsgBox("The report styles have been refreshed")
                Catch ex As Exception

                End Try
                '                
                rng.Select()
                '
            Case "grpStylesTools_resetStyleShort"
                rng = objGlobals.glb_get_wrdSel.Range
                Try
                    objRptMgr.Rpt_Styles_resetStyles_fromTemplate()
                    MsgBox("The report styles have been refreshed")
                Catch ex As Exception

                End Try
                '                
                rng.Select()
                '
            Case "grpStylesTools_resetStyleLandscape"
                rng = objGlobals.glb_get_wrdSel().Range
                Try
                    objRptMgr.Rpt_Styles_resetStyles_fromTemplate()
                    MsgBox("The report styles have been refreshed")
                Catch ex As Exception

                End Try
                '                
                rng.Select()
                '
            Case "grpStylesTools_to_PrintDefault"
                myDoc = objGlobals.glb_get_wrdActiveDoc()
                objPrint = New cPrintAndDisplayServices()
                objPrint.colour_display_ToPrintMode(myDoc)

            Case "grpStylesTools_to_DisplayDefault"
                myDoc = objGlobals.glb_get_wrdActiveDoc()
                objPrint = New cPrintAndDisplayServices()
                objPrint.colour_display_ToDefault(myDoc)

        End Select
        '
finis:
    End Sub

    Private Sub PIF_PgS_grpWCAGTools_Click(sender As Object, e As RibbonControlEventArgs) Handles grpRbn_Mgmnt_removeRbn.Click, grpWCAG_tool_convertAllStyles_toBlack.Click, grpWCAG_tool_tableHeaderColour_all.Click
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objWCAGMgr As New cWCAGMgr()
        Dim objGlobals As New cGlobals()
        Dim objFileMgr As New cFileHandler()
        Dim objMsgMgr As cMessageManager
        Dim objProp As New cPropertyMgr()
        Dim myDoc As Word.Document
        Dim myDocInfo As System.IO.FileInfo
        Dim strDocs As String()
        Dim strNewFileName, strColourPickerMode As String
        Dim styl As Word.Style
        Dim isOK As Boolean
        'Dim frmPicker As frm_colorPicker
        '
        myDoc = objGlobals.glb_get_wrdActiveDoc()
        isOK = True
        '
        objGlobals.glb_cursors_setToWait()

        objGlobals.glb_get_wrdApp.ScreenUpdating = False
        '
        Try
            Select Case e.Control.Id
                Case "grpWCAG_tool_rbnDelete"
                    objWCAGMgr.wcag_rbn_del(myDoc)
                Case "grpRbn_Mgmnt_removeRbn"
                    'MsgBox("REmove Rbn")
                    strNewFileName = myDoc.FullName
                    '
                    If myDoc IsNot Nothing Then
                        Try
                            ' Ensure the document is saved before proceeding
                            If Not (myDoc.Path = "") Then
                                Dim userResponse = MsgBox("This document needs to be saved before the ribbon can be removed." & vbCrLf & "Do you want to save it now?", vbYesNo)
                                If userResponse = vbYes Then
                                    objFileMgr.file_get_saveTimeStampedCopy(myDoc, "")
                                    'myDoc.Save()
                                Else
                                    Exit Select
                                End If
                            End If
                            ' Proceed with ribbon removal
                            '
                            objProp.prps_del_customProperty("_AssemblyLocation", myDoc)
                            objProp.prps_del_customProperty("_AssemblyName", myDoc)
                            '
                            myDoc.AttachedTemplate = objGlobals.glb_getTmpl_FullName()
                            '
                            myDoc.Save()

                            MsgBox("The ribbon has been removed and the document saved as " + vbCrLf + vbCrLf + myDoc.FullName + vbCrLf + vbCrLf + "You'll need to close and re-open the document to make the new ribbon setting 'stick'")
                        Catch ex2 As Exception
                            MsgBox("Error removing Ribbon: " & ex2.Message)
                        End Try
                    End If
                    '
                Case "grpRbn_Mgmnt_addRbn"
                    'MsgBox("Add Rbn")
                    strDocs = objFileMgr.getDocsToImport("Select Document",, False)
                    myDocInfo = New System.IO.FileInfo(strDocs(0))
                    '
                    '** Adjust for OneDrive.. See above
                    If myDocInfo.Exists Then
                        If Not objFileMgr.isFileOpen(New System.IO.FileInfo(strDocs(0))) Then
                            myDoc = objGlobals.glb_get_wrdApp.Documents.Open(strDocs(0))
                        Else
                            myDoc = objGlobals.glb_get_wrdApp.Documents.Item(strDocs(0))
                        End If
                        '
                        strNewFileName = objFileMgr.file_get_newFileName(myDoc, myDocInfo.DirectoryName, "")
                        '
                        'objWCAGMgr.wcag_rbn_addAAC(myDoc, "testMachine")
                        objWCAGMgr.wcag_rbn_addAAC(myDoc, "generalReport")
                        'objWCAGMgr.wcag_rbn_addAAC(myDoc, "mikl.net.au")
                        'objWCAGMgr.wcag_rbn_addAAC(myDoc, "generalReport_Internal")
                        '
                        myDoc.Save()
                        myDoc.Saved = True
                        '
                        Try
                            myDoc.Close(WdSaveOptions.wdSaveChanges)
                            'Can't close the document because the code to close it comes from the document?
                            'myDoc = objGlobals.glb_get_wrdApp.Documents.Open(strDocs(0))
                            '
                        Catch ex_grpRbn_1p0 As Exception

                        End Try
                        '
                        MsgBox("The 'AAC' ribbon has been restored." + vbCr + vbCr + "The document has been saved and closed in order to make the new ribbon setting stick.")

                    Else
                        MsgBox("This document needs to be saved before the ribbon can be restored")
                    End If

                Case "grpWCAG_tool_tocUnlink"
                    objWCAGMgr.wcag_toc_unlinkFields(myDoc)
                    '
                    '
                Case "grpWCAG_tool_convertAllStyles", "grpWCAG_tool_convertAllStyles_toBlack"
                    objMsgMgr = New cMessageManager()
                    If objMsgMgr.msgMgr_dlg_stylesToBlack() Then
                        objWCAGMgr.wcag_styles_setForWCAG(myDoc)
                        MsgBox("The colour change to the styles is complete")
                    End If
                Case "grpWCAG_tool_tableHeaderColour_all"
                    objMsgMgr = New cMessageManager()
                    strColourPickerMode = "tbl_Header_Colour_all"
                    If objMsgMgr.msgMgr_dlg_fillAllTableHeaders() Then
                        objGlobals.glb_show_ColorPicker(strColourPickerMode)
                        'frmPicker = New frm_colorPicker(strColourPickerMode)
                        'frmPicker.Show()
                    Else

                    End If

                    '
                Case "grpWCAG_tool_convertBackColourPages"
                    objWCAGMgr.wcag_convert_backColour(myDoc)
                Case "grpWCAG_tool_convertBanners"
                    objWCAGMgr.wcag_convert_bannersToWCAG(myDoc)
                Case "grpWCAG_tool_convertHeadersFooters"
                    'We'll do the Header-Company Name style here just so that the Header/Footer
                    'works.. It saves on the time taken to adjust all styles
                    styl = myDoc.Styles.Item("Header-Company Name")
                    styl.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
                    '
                    'sect = objGlobals.glb_get_wrdSect()
                    'objHFMgr.hf_hfs_convertToText(sect, "header")

                    objWCAGMgr.wcag_convert_headersToWCAG(myDoc)
                    objWCAGMgr.wcag_convert_footersToWCAG(myDoc)
                Case "grpWCAG_tool_unLinkAllFields"
                    myDoc.Fields.Unlink()
            End Select
        Catch ex As Exception
            isOK = False
        End Try
        objGlobals.glb_cursors_setToNormal()

        '
        objGlobals.glb_get_wrdApp.ScreenUpdating = True
        objGlobals.glb_get_wrdApp.ScreenRefresh()
        '
        If isOK Then
            'MsgBox("Conversion complete")
        Else
            'MsgBox(control.Id + " conversion failed")
        End If
        '
    End Sub
    '
    Private Sub PIF_PgS_grpReport_Click(sender As Object, e As RibbonControlEventArgs) Handles grpRpt_mnu_btn_NewChapter_inFront_bblk.Click, grpRpt_mnu_btn_NewChapter_behind_bblk.Click, grpReprt_btn_buildLandscapeReport.Click, grpReport_btn_buildPortraitReport.Click, grpReport_btn_buildAABrief.Click, grpReport_mnu01_SelectedText.Click, grpReport_mnu01_SelectedTblCells.Click, grpReport_mnu_CaseStudies_RecolourLogo_toWhite.Click, grpReport_mnu_CaseStudies_RecolourLogo_Reset.Click, grpReport_mnu_CaseStudies_RecolourFooter_toWhite.Click, grpReport_mnu_CaseStudies_RecolourFooter_Reset.Click, grpReport_mnu_CaseStudies_HalfPage.Click, grpReport_mnu_CaseStudies_FullPage.Click, grpReport_mnu_CaseStudies_CaseStudyHeading.Click, tabPgs_grpRpt_btn_buildPrtReport_sw.Click, tabPgs_grpRpt_btn_buildLndReport_sw.Click, tabPgs_grpRpt_btn_buildBrfReport_sw.Click, grpReport_btn_newDivider_Chpt_bblk.Click, grpReport_btn_newDivider_Chpt.Click, grpRpt_mnu_btn_NewChapter_inFront.Click, grpRpt_mnu_btn_NewChapter_behind.Click, grpReport_PlH_convertToInline_findAllFloatingTables_2.Click, grpReport_PlH_convertToInline_findAllFloatingTables.Click, grpReport_tbHome_btn_buildPortraitReport.Click, grpReport_tbHome_btn_buildLandscapeReport.Click, grpReport_tbHome_btn_buildAABrief.Click
        Dim objSectMgr As New cSectionMgr()
        Dim objChpt As New cChptBody()
        Dim objViewMgr As New cViewManager()
        Dim objPropsMgr As New cPropertyMgr()
        Dim objDivMgr As New cChptDivider()
        Dim objRptMgr As New cReport()
        Dim objpgNumMgr As New cPageNumberMgr()
        Dim objMsgMgr As New cMessageManager()
        Dim objStylesMgr As New cStylesManager()
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objWrkAround As New cWorkArounds()
        Dim objBBlkMgr As New cBBlocksHandler()
        Dim objContactsPgMgr As New cContactsMgr()
        '
        Dim objIsOKMgr As New cIsOKToDo()
        Dim objLogoMgr As New cLogosMgr()
        Dim objPlhBase As New cPlHBase()
        Dim objCaseStudyMgr As cCaseStudyMgr
        '
        'Dim frmPicker As frm_colorPicker
        'Dim frmPicker As frm_colorPicker02
        '
        'Dim objCpMgr As New cCoverPageMgr()
        'Dim objContactsPgMgr As New cContactsMgr()
        Dim objTOCMgr As New cTOCMgr()
        '
        Dim frm_doc_Placeholders As frm_findTables
        Dim objGlobals As New cGlobals()
        Dim rng As Word.Range
        '
        Dim tbl As Word.Table
        Dim sect As Word.Section
        Dim placeBehind As Boolean
        Dim strColourPickerMode As String
        Dim strHeading As String
        Dim strRptMode As String
        '
        Dim strErrorMsg, strMsg As String
        '
        strMsg = ""
        strErrorMsg = ""
        placeBehind = True
        sect = objSectMgr.objGlobals.glb_get_wrdApp.Selection.Sections.Item(1)
        '
        objGlobals.glb_cursors_setToWait()
        '
        'Refresh and turn off updating
        objSectMgr.objGlobals.glb_screen_update(False)

        '
        strRptMode = objRptMgr.Rpt_Mode_Get()

        Select Case e.Control.Id
            Case "grpReport_btn_buildStylesGuide"
                objRptMgr = New cReport()
                '
                If objMsgMgr.deleteAllMessage Then
                    strRptMode = objRptMgr.Rpt_Mode_SetAs_Std()
                    '
                    'objRptMgr.Rpt_build_newReport_PrtandLnd()                      'Software build
                    objRptMgr.Rpt_build_fastReportOrBrief_byCopy("stylesGuide")          'Build from embedded example
                Else

                End If
            Case "grpReport_btn_buildStylesGuide_Accessible"
                objRptMgr = New cReport()
                '
                If objMsgMgr.deleteAllMessage Then
                    strRptMode = objRptMgr.Rpt_Mode_SetAs_Std()
                    '
                    'objRptMgr.Rpt_build_newReport_PrtandLnd()                      'Software build
                    objRptMgr.Rpt_build_fastReportOrBrief_byCopy("stylesGuide_AccessibleAware")          'Build from embedded example
                Else

                End If


            Case "grpReport_btn_buildPortraitReport", "grpReport_tbHome_btn_buildPortraitReport"
                objRptMgr = New cReport()
                objRptMgr.Rpt_build_fastReportOrBrief_fromTemplate()
                '
                'If objMsgMgr.deleteAllMessage Then
                'strRptMode = objRptMgr.Rpt_Mode_SetAs_Std()
                '
                'objRptMgr.Rpt_build_newReport_PrtandLnd()              'Software build
                'objRptMgr.Rpt_build_fastReportOrBrief_byCopy()          'Build from embedded example
                'Else

                'End If
                '
            Case "tabPgs_grpRpt_btn_buildPrtReport_sw"
                objRptMgr = New cReport()
                '
                If objMsgMgr.deleteAllMessage Then
                    strRptMode = objRptMgr.Rpt_Mode_SetAs_Std()
                    '
                    objRptMgr.Rpt_build_newReport_PrtandLnd()              'Software build
                    'objRptMgr.Rpt_build_NewReport_PrtandLnd_Copy()          'Build from embedded example
                Else

                End If
                '
            Case "grpReprt_btn_buildLandscapeReport", "grpReport_tbHome_btn_buildLandscapeReport"
                objRptMgr = New cReport()
                objRptMgr.Rpt_build_fastReportOrBrief_fromTemplate("Lnd")

                '
                'If objMsgMgr.deleteAllMessage Then
                'strRptMode = objRptMgr.Rpt_Mode_SetAsLandScape()
                '
                'objRptMgr.Rpt_build_newReport_PrtandLnd()              'Software build
                'objRptMgr.Rpt_build_fastReportOrBrief_byCopy("Lnd")     'Build from embedded example

                'objpgNumMgr.pgNum_setBody_numFormat(objSectMgr.objGlobals.glb_get_wrdActiveDoc(), "std")
                'Else

                'End If
                '
            Case "tabPgs_grpRpt_btn_buildLndReport_sw"
                objRptMgr = New cReport()
                '
                If objMsgMgr.deleteAllMessage Then
                    strRptMode = objRptMgr.Rpt_Mode_SetAsLandScape()
                    '
                    objRptMgr.Rpt_build_newReport_PrtandLnd()              'Software build
                    'objRptMgr.Rpt_build_NewReport_PrtandLnd_Copy("Lnd")     'Build from embedded example

                    'objpgNumMgr.pgNum_setBody_numFormat(objSectMgr.objGlobals.glb_get_wrdActiveDoc(), "std")
                Else

                End If
                '
            Case "grpReport_btn_buildAABrief", "grpLetter_buildAABrief", "grpReport_tbHome_btn_buildAABrief"
                objRptMgr = New cReport()
                objRptMgr.Rpt_build_fastReportOrBrief_fromTemplate("Brf")

                objWrkAround.wrk_fix_forCursorRace()
                '
            Case "tabPgs_grpRpt_btn_buildBrfReport_sw"
                objRptMgr = New cReport()
                '
                If objMsgMgr.deleteAllMessage Then
                    'strRptMode = objRptMgr.Rpt_Mode_SetAsShort()
                    strRptMode = objRptMgr.Rpt_Mode_SetAsAABrief()
                    objPropsMgr.prps_del_customProperty("pgNumberFormat", objGlobals.glb_get_wrdActiveDoc())
                    '
                    objRptMgr.Rpt_build_newAABrief_PrtandLnd()             'Software build
                    'objRptMgr.Rpt_build_NewReport_PrtandLnd_Copy("Brf")     'Build from embedded example

                Else

                End If
                '
                objWrkAround.wrk_fix_forCursorRace()
                '
            Case "grpReport_btn_newDivider_Chpt_bblk"
                strErrorMsg = objIsOKMgr.isOKto_Insert_DividerChpt(sect)
                '
                'If objGlobals.glb_selection_IsInTable() Then strErrorMsg = "Selection in table"
                '
                If strErrorMsg = "" Then
                    sect = objGlobals.glb_get_wrdSect()
                    rng = sect.Range
                    rng = objBBlkMgr.bblk_insert_sectionInFront("aa_Divider_Std")
                    rng = objDivMgr.chptBase_getRange_Heading1(rng.Sections.First)
                    rng.Select()
                    '
                Else
                    MsgBox(strErrorMsg)
                End If
                '
                '
            Case "grpReport_newDivider_Chpt", "grpReport_btn_newDivider_Chpt"
                strErrorMsg = objIsOKMgr.isOKto_Insert_DividerChpt(sect)
                '
                Select Case strRptMode
                    Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                        Try
                            If strErrorMsg = "" Then

                                objDivMgr = New cChptDivider()
                                strHeading = "Divider Title" + vbCr + "Sub Heading"
                                'sect = objDivMgr.div_insert_newBody(Not placeBehind, sect, strHeading)
                                sect = objDivMgr.div_insert_newBody(Not placeBehind, sect)
                                objTOCMgr.toc_update_TOCs(sect.Range.Document)
                                '
                                rng = objDivMgr.chptBase_getRange_Heading1(sect)
                                rng.Select()
                            Else
                                MsgBox(strErrorMsg)
                            End If

                        Catch ex As Exception
                            MsgBox("Error In Divider Insert")
                        End Try
                        '
                    Case objRptMgr.rpt_isBrief
                        MsgBox(objMsgMgr.msgMgr_msg_notAvailableInBrief())

                End Select
                '
                '
            Case "grpRpt_mnu_btn_NewChapter_inFront_bblk"
                strErrorMsg = objIsOKMgr.isOKto_Insert_ChapterInFront(sect, strRptMode)
                objViewMgr.vw_change_ColumnsAndRows(sect)
                '
                If strErrorMsg = "" Then
                    rng = objBBlkMgr.bblk_insert_sectionInFront("aa_Chpt_Std")
                    rng.Select()
                    '
                Else
                    MsgBox(strErrorMsg)
                End If
                '
            Case "grpReport_newChapter", "grpRpt_mnu_btn_NewChapter_inFront"
                strErrorMsg = objIsOKMgr.isOKto_Insert_ChapterInFront(sect, strRptMode)
                objViewMgr.vw_change_ColumnsAndRows(sect)
                '
                objSectMgr.objGlobals.glb_screen_update(False)

                Try
                    If strErrorMsg = "" Then
                        rng = objChpt.chpt_insert_Body(Not placeBehind, sect, strRptMode)
                        '
                        'Must override the orignal setup defined in cChptBnr.bnr_get_BannerSettings which
                        'rng = sect.Range
                        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
                        'rng.Paragraphs.Item(1).Range.Delete()
                        'rng = objStylesMgr.styles_insert_StartupText_ReportBody(rng)
                        '
                        '
                    Else
                        MsgBox(strErrorMsg)
                    End If
                Catch ex As Exception
                    MsgBox("Error In Chapter Insert")
                End Try
                '
                'objViewMgr.vw_change_ColumnsAndRows(sect)
                '
            Case "grpRpt_mnu_btn_NewChapter_behind_bblk"
                'To insert behind, go to the next section and insert in front
                strErrorMsg = objIsOKMgr.isOKto_Insert_ChapterBehind(sect, strRptMode)
                objViewMgr.vw_change_ColumnsAndRows(sect)
                '
                Try
                    If strErrorMsg = "" Then
                        Dim myDoc As Word.Document
                        myDoc = sect.Range.Document
                        sect = objGlobals.glb_get_wrdSect
                        If Not (sect.Index = myDoc.Sections.Last.Index) Then
                            sect = myDoc.Sections.Item(sect.Index + 1)
                            rng = sect.Range
                            rng.Collapse(WdCollapseDirection.wdCollapseStart)
                            rng.Select()
                            rng = objBBlkMgr.bblk_insert_sectionInFront("aa_Chpt_Std")
                            rng.Select()
                        Else
                            MsgBox("Inserting behind the last section is not permitted")
                        End If
                        '
                    Else
                        MsgBox(strErrorMsg)
                    End If

                Catch ex As Exception

                End Try

            Case "grpReport_newChapter_Behind", "grpRpt_mnu_btn_NewChapter_behind"
                '
                strErrorMsg = objIsOKMgr.isOKto_Insert_ChapterBehind(sect, strRptMode)
                '
                'strErrorMsg = ""
                Try
                    If strErrorMsg = "" Then
                        rng = objChpt.chpt_insert_Body(placeBehind, sect, strRptMode)
                        '
                        rng = objChpt.chptBase_getRange_Heading1(sect)
                        rng.Select()
                    Else
                        MsgBox(strErrorMsg)
                    End If
                    '
                Catch ex As Exception
                    MsgBox("Error In Chapter Insert")
                End Try
                '

            Case "grpReport_buildNew_Short"
                objRptMgr = New cReport()
                strRptMode = objRptMgr.Rpt_Mode_SetAsShort()
                '
                objRptMgr.Rpt_build_newReport_Short()
                objpgNumMgr.pgNum_setBody_numFormat(objSectMgr.objGlobals.glb_get_wrdActiveDoc(), "std")

            Case "grpReport_Long"
                'If it is a Short Portrait Report, then we can convert it to the Long version.
                'Note that "Rpt_ModeChange_ToShort()" calls "Rpt_Mode_SetAsLong()" which will chnage
                'any styles that need tpo be changed (typically H1 style)
                '
                '
                objRptMgr.Rpt_Mode_SetAs_Std()
                objGlobals.glb_get_wrdActiveDoc.Styles.Item("Heading 1").ParagraphFormat.PageBreakBefore = True
                '
                'May need to redraw the cover page icons
                'Me.ribbon.InvalidateControl("gal_CoverPages")
                'oldRange.Select()
            Case "grpReport_Short"
                'If it is a Standard Portrait Report, then we can convert it to the Short version.
                'Note that "Rpt_ModeChange_ToShort()" calls "Rpt_Mode_SetAsShort()" which will chnage
                'any styles that need tpo be changed (typically H1 style)
                '
                objRptMgr.Rpt_Mode_SetAsShort()
                objGlobals.glb_get_wrdActiveDoc.Styles.Item("Heading 1").ParagraphFormat.PageBreakBefore = False
                '
                'May need to redraw the cover page icons
                'Me.ribbon.InvalidateControl("gal_CoverPages")
                'oldRange.Select()

            Case "grpReport_PlH_convertToInline_findAllFloatingTables", "grpReport_PlH_convertToInline_findAllFloatingTables_2"
                frm_doc_Placeholders = New frm_findTables()
                '
                frm_doc_Placeholders.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
                frm_doc_Placeholders.TopMost = True
                frm_doc_Placeholders.Show()
                '
            Case "grpReport_mnu_CaseStudies_FullPage"
                'Inserts a bounded section at the current selection. When finished the selection is at the 
                'beginning of the nhew bounded section

                strErrorMsg = objIsOKMgr.isOKto_doAction_inReportBody()
                'objHFhf_hfs_getHfTableEdges
                '
                If strErrorMsg = objIsOKMgr._isOK Then
                    '
                    objCaseStudyMgr = New cCaseStudyMgr()
                    strErrorMsg = objCaseStudyMgr.cst_insert_fullPageCaseStudy(objSectMgr.objGlobals.glb_get_wrdSelRng)
                    '
                    Select Case strErrorMsg
                        Case "inTable"
                            objMsgMgr.msg_insertionPoint_IsIn_Or_JustUnderATable()
                            '
                        Case ""

                    End Select
                Else
                    MsgBox(strErrorMsg)
                End If
                                '
            Case "grpReport_mnu_CaseStudies_HalfPage"
                'Inserts a bounded section at the current selection. When finished the selection is at the 
                'beginning of the nhew bounded section
                objCaseStudyMgr = New cCaseStudyMgr()
                rng = objGlobals.glb_get_wrdSelRng()
                '
                tbl = objCaseStudyMgr.cst_insert_partialPageCaseStudy(rng)

                'tbl = objCaseStudyMgr.Plh_insert_PlaceHolder_WithTest(rng, "CaseStudy_HalfPage")
                'rng = tbl.Range.Cells.Item(1).Range
                'rng.Text = "Insert Case study text here"
                rng.Select()

               ' tbl = objPlhBase.Plh_insert_PlaceHolder_WithTest(rng, "CaseStudy_HalfPage")

                '
            Case "grpReport_mnu_CaseStudies_CaseStudyHeading"
                'Will apply the Case Study Heading to the selected paragraphs
                Try
                    objCaseStudyMgr = New cCaseStudyMgr()
                    objCaseStudyMgr.cst_insert_Caption(objGlobals.glb_get_wrdSelRngAll)

                    'objStylesMgr.applyStyleToSelection("Heading (CaseStudy)")

                Catch ex As Exception
                    MsgBox("Failed To apply 'Heading (CaseStudy)'")
                End Try

            Case "grpReport_mnu_CaseStudies_RecolourLogo_toWhite"
                'Will recolour the Acil Allen logo
                Try
                    objLogoMgr.logos_set_colour(sect, RGB(255, 255, 255), -1)

                Catch ex As Exception
                    MsgBox("Failed to recolour the logo")
                End Try
                '
            Case "grpReport_mnu_CaseStudies_RecolourLogo_Reset"
                'Will recolour the Acil Allen logo
                Try
                    objLogoMgr.logos_set_colour(sect, RGB(0, 0, 0), -1)

                Catch ex As Exception
                    MsgBox("Failed to reset the logo colour")
                End Try
                '
            Case "grpReport_mnu_CaseStudies_RecolourFooter_toWhite"
                objHfMgr.hf_set_textColourFooter(objGlobals.glb_get_wrdSect, RGB(255, 255, 255), "newColour", "primary")

            Case "grpReport_mnu_CaseStudies_RecolourFooter_Reset"
                objHfMgr.hf_set_textColourFooter(objGlobals.glb_get_wrdSect, RGB(255, 255, 255), "resetColour", "primary")

            Case "grpReport_mnu01_SelectedText", "grpReport_mnu01_SelectedTblCells"
                strColourPickerMode = "text_Colour"
                Select Case e.Control.Id
                    Case "grpTbls_setTableTextCustomColour"
                        strColourPickerMode = "text_Colour"
                    Case "grpReport_mnu01_SelectedTblCells"
                        strColourPickerMode = "tbl_Cells"
                End Select
                '
                objGlobals.glb_show_ColorPicker(strColourPickerMode)
                '
        End Select
        '
        objSectMgr.objGlobals.glb_screen_update(True)
        objGlobals.glb_cursors_setToNormal()
        '
        objWrkAround.wrk_fix_forCursorRace()
        'objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = True

    End Sub
    '

    Private Sub PIF_PgS_grpExecSum_Click(sender As Object, e As RibbonControlEventArgs) Handles grpExecSum_ExecSum_Grey_bblk.Click, grpExecSum_ExecSum_bblk.Click, grpExecSum_ExecSum_Grey.Click, grpExecSum_ExecSum.Click
        Dim objSectMgr As New cSectionMgr()
        Dim objChpt As New cChptExec()
        Dim objRptMgr As New cReport()
        Dim objIsOKMgr As New cIsOKToDo()
        Dim objBackPanelMgr As New cBackPanelMgr()
        Dim objGlobals As New cGlobals()
        Dim objMsgMgr As New cMessageManager()
        Dim objWrkAround As New cWorkArounds()
        Dim objBBlkMgr As New cBBlocksHandler()
        Dim objChptBase As New cChptBase()
        Dim sect As Word.Section
        Dim rng As Word.Range
        Dim hf As Word.HeaderFooter
        Dim tbl As Word.Table
        Dim placebehind As Boolean
        Dim strErrorMsg, strRptMode As String
        '
        strErrorMsg = ""
        sect = objSectMgr.objGlobals.glb_get_wrdSect()
        objSectMgr.objGlobals.glb_screen_update(False)

        placebehind = False

        Select Case e.Control.Id
            Case "grpExecSum_ExecSum"
                strErrorMsg = objIsOKMgr.isOKto_Insert_ESummary(sect)
                '
                If strErrorMsg = "" Then
                    tbl = objChpt.es_insert_section(placebehind, sect, objRptMgr.Rpt_Mode_Get())

                Else
                    MsgBox(strErrorMsg)
                End If
                '
            Case "grpExecSum_ExecSum_Grey"
                strRptMode = objRptMgr.Rpt_Mode_Get()
                '
                Select Case strRptMode
                    Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                        strErrorMsg = objIsOKMgr.isOKto_Insert_ESummary(sect)
                        If strErrorMsg = "" Then
                            objChpt.es_insert_section(placebehind, sect, strRptMode)
                            hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                            objBackPanelMgr.pnl_BackPanel_Insert(hf, objGlobals._glb_colour_CaseStudy_Grey)
                            'rng = objChpt.chptBase_getRange_Heading1(sect)
                            '
                            'rng.Select()
                        Else
                            MsgBox(strErrorMsg)
                        End If
                        '
                    Case objRptMgr.rpt_isBrief
                        MsgBox(objMsgMgr.msgMgr_msg_notAvailableInBrief())

                End Select
                '
            Case "grpExecSum_ExecSum_bblk", "grpExecSum_ExecSum_Grey_bblk"
                strErrorMsg = objIsOKMgr.isOKto_Insert_ESummary(sect)
                '
                If strErrorMsg = "" Then
                    rng = sect.Range
                    rng = objBBlkMgr.bblk_insert_sectionInFront("aa_Chpt_ES")
                    Select Case e.Control.Id
                        Case "grpExecSum_ExecSum_bblk"
                            rng = objChptBase.chptBase_getRange_Heading1(rng.Sections.First)
                        Case "grpExecSum_ExecSum_Grey_bblk"
                            sect = rng.Sections.First
                            hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                            objBackPanelMgr.pnl_BackPanel_Insert(hf, objGlobals._glb_colour_CaseStudy_Grey)
                            rng = objChptBase.chptBase_getRange_Heading1(rng.Sections.First)
                    End Select
                    rng.Select()

                Else
                    MsgBox(strErrorMsg)
                End If
                '
                'objChpt.chptBase_select_Chapter(tbl)
            Case "grpExecSum_ExecSum_Behind"
                'tbl = objChpt.es_insert_section(placeBehind, sect, objRptMgr.Rpt_Mode_Get())
                'objChpt.chptBase_select_Chapter(tbl)
            Case "grpExecSum_Landscape"
                sect = objChpt.es_insert_LndBounded()
                rng = sect.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Select()
        End Select
        '
        objWrkAround.wrk_fix_forCursorRace()
        '
        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = True

    End Sub
    '
    Private Sub PIF_Plh_grpEquations_Click(sender As Object, e As RibbonControlEventArgs) Handles grpEquations_Numbered.Click
        'Dim oMaths As New OMaths
        Dim objTblsMgr As New cTablesMgr()
        Dim objMaths As New cMaths()
        Dim objPlHMgr As New cPlHBase()
        Dim objGraphicsMgr As New cGraphicsMgr()
        Dim sect As Word.Section
        Dim strMsg As String
        Dim rng As Word.Range
        Dim tbl As Word.Table
        '
        objTblsMgr.glb_get_wrdApp.ScreenUpdating = False
        '
        rng = objTblsMgr.glb_get_wrdSelRng()
        sect = objTblsMgr.glb_get_wrdSect()
        '
        Select Case e.Control.Id
            Case "grpEquations_Numbered"
                'objPlHMgr.inse
                ' objGraphicsMgr.test()
                'GoTo finis
                'Call m_EquationManager.InsertNumberedEquation(control)
                strMsg = objPlHMgr.Plh_is_OKToInsert(sect, True)
                '
                If strMsg = "" Then
                    tbl = objMaths.mth_equationEditor_insert(objTblsMgr.glb_get_wrdSelRng())
                Else
                    MsgBox(strMsg)
                End If
                '
            Case "Gallery1"
                objTblsMgr.glb_get_wrdApp._CommandBars.ExecuteMso("InsertEquationGallery")
            Case Else
        End Select
        '
finis:
        '
        objTblsMgr.glb_get_wrdApp.ScreenUpdating = True
        '
    End Sub
    '
    Private Sub PIF_Plh_grpFloatMgmnt_Click(sender As Object, e As RibbonControlEventArgs) Handles grpReport_PlHFloat_lock_toMarginsLeftAndTop.Click, grpReport_PlH_LockToTop.Click, grpReport_PlH_LockToParagraphAndColumn.Click, grpReport_PlH_LockToParagraph.Click, grpReport_PlH_convertToInline.Click, grpReport_PlH_FloatWide.Click, grpReport_PlH_FloatEdgeToEdge.Click, grpReport_PlH_TwoColumnWidth.Click, grpReport_PlH_ColumnWidth.Click, grpReport_PlH_FloatMarginToMargin.Click
        Dim objSectMgr As New cSectionMgr()
        Dim objFloatMgr As New cPlHFloatingMgr()
        Dim objisOK As New cIsOKToDo()
        Dim objWrkAround As New cWorkArounds()
        Dim sect As Word.Section
        Dim tbl As Word.Table
        Dim strMsg, strMsg2, strMsg3, strMsgIsOK As String
        '
        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = False
        '
        strMsg = "Please make certain that your cursor is in the 'PlaceHolder', or Banner that you wish to float"
        strMsg2 = "Please make certain that your cursor is in the floating 'PlaceHolder' that you need to make 'inline'"
        strMsg3 = "Please make certain that your cursor is in the floating 'PlaceHolder' (table, figure or box) that you wish to make inline"
        '
        strMsgIsOK = objisOK.isOKto_doAction_inReportBody()
        '
        If strMsgIsOK = objisOK._isOK Then
            Try
                'strRptMode = objRptMgr.Rpt_Mode_Get()
                sect = objSectMgr.objGlobals.glb_get_wrdSelRng.Sections.Item(1)
                tbl = objSectMgr.objGlobals.glb_get_wrdSelTbl()
                '
                Select Case e.Control.Id
                    Case "grpReport_PlH_convertToInline"
                        If Not IsNothing(tbl) Then
                            objFloatMgr.PlHFloat_convert_toInLine(tbl)
                            MsgBox("The selected table is now 'inline'." + vbCrLf + vbCrLf + "If you wish to reset it to floating, select one of the options from the 'PlaceHolder Mgmnt' menu in this group")
                        Else
                            MsgBox(strMsg3)
                        End If

                    Case "grpReport_PlH_LockToTop"
                        If Not IsNothing(tbl) Then
                            objFloatMgr.PlHFloat_lock_toMarginsTop(tbl)
                        Else
                            MsgBox(strMsg)
                        End If
                    Case "grpReport_PlHFloat_lock_toMarginsLeftAndTop"
                        If Not IsNothing(tbl) Then
                            objFloatMgr.PlHFloat_lock_toMarginsLeftAndTop(tbl)
                        Else
                            MsgBox(strMsg)
                        End If
                    Case "grpReport_PlH_LockToParagraph"
                        If Not IsNothing(tbl) Then
                            objFloatMgr.PlHFloat_lock_toParagraphAndMarginLeft(tbl)
                        Else
                            MsgBox(strMsg)
                        End If

                    Case "grpReport_PlH_LockToParagraphAndColumn"
                        If Not IsNothing(tbl) Then
                            objFloatMgr.PlHFloat_lock_toParagraphAndColumn(tbl)
                        Else
                            MsgBox(strMsg)
                        End If

                    '
                    Case "grpReport_PlH_FloatEdgeToEdge"
                        '
                        '
                        If Not IsNothing(tbl) Then
                            strMsg = objFloatMgr.Plh_Float_EdgToEdge(tbl)
                            'strMsg = objFloatMgr.Plh_Float_Wide(tbl)
                            If strMsg <> "" Then
                                MsgBox(strMsg)
                            End If
                        Else
                            MsgBox(strMsg2)
                        End If
                    '                                        '
                    Case "grpReport_PlH_FloatWide"
                        '
                        '
                        If Not IsNothing(tbl) Then
                            strMsg = objFloatMgr.Plh_Float_Wide(tbl)
                            If strMsg <> "" Then
                                MsgBox(strMsg)
                            End If
                        Else
                            MsgBox(strMsg2)
                        End If
                    '
                    Case "grpReport_PlH_FloatMarginToMargin"
                        '
                        If Not IsNothing(tbl) Then
                            strMsg = objFloatMgr.Plh_Float_MarginToMargin(tbl)
                            If strMsg <> "" Then
                                MsgBox(strMsg)
                            End If
                        Else
                            MsgBox(strMsg2)
                        End If
                    '
                    Case "grpReport_PlH_ColumnWidth"
                        '
                        If Not IsNothing(tbl) Then
                            strMsg = objFloatMgr.Plh_Float_FullColumn(tbl)
                            If strMsg <> "" Then
                                MsgBox(strMsg)
                            End If
                        Else
                            MsgBox(strMsg2)
                        End If
                        '
                    Case "grpReport_PlH_TwoColumnWidth"
                        If Not IsNothing(tbl) Then
                            strMsg = objFloatMgr.Plh_Float_FullColumn_x2(tbl)
                            If strMsg <> "" Then
                                MsgBox(strMsg)
                            End If
                        Else
                            MsgBox(strMsg2)
                        End If

                End Select
                '
            Catch ex As Exception
                '
                objSectMgr.objGlobals.glb_screen_update(True)
                '
            End Try

        Else
            MsgBox(strMsgIsOK)
        End If
        '
        objWrkAround.wrk_fix_forCursorRace()
        objSectMgr.objGlobals.glb_screen_update(True)
        '
        '
    End Sub
    '
    Private Sub PIF_Plh_grpPicts_Click(sender As Object, e As RibbonControlEventArgs) Handles grpPicts_PasteAsPic.Click
        Dim objImgMgr As New cImageMgr()
        Dim objGlobals As New cGlobals()
        '
        objGlobals.glb_get_wrdApp.ScreenUpdating = False
        '
        Select Case e.Control.Id
            Case "grpPicts_PasteAsPic"
                Call objImgMgr.PasteAsPicture()
            Case Else
        End Select
        '
        objGlobals.glb_get_wrdApp.ScreenUpdating = True
        '
    End Sub
    '
    Private Sub PIF_Plh_grpTables_Click(sender As Object, e As RibbonControlEventArgs) Handles grpPlh_btn_buildCustomTable.Click, grpTbls_setTableTextCustomColour.Click, grpTbls_fillCellsWithCustomColour.Click, grpTbls_StyleSet_TableListNumbers.Click
        Dim frm As frm_TableBuilder
        Dim objGlobals As New cGlobals()
        Dim objStylesMgr As New cStylesManager
        Dim objPlHTable As New cPlHTable()
        Dim objTools As New cTools()
        Dim objFlds As New cFieldsMgr()
        Dim objMsgMgr As New cMessageManager()
        Dim strColourPickerMode As String
        'Dim frmPicker As frm_colorPicker
        Dim rng As Word.Range
        Dim objTables As New cTablesMgr()
        Dim tbl As Word.Table
        'Dim dlg_InsertTable As Word.Dialog = Globals.ThisAddin.Application.Dialogs(Word.WdWordDialog.wdDialogTableInsertTable)
        '
        objGlobals.glb_get_wrdApp.ScreenUpdating = False
        '
        Try
            Select Case e.Control.Id
                Case "grpPlh_btn_buildCustomTable"
                    rng = objGlobals.glb_get_wrdSelRng()
                    If rng.Tables.Count <> 0 Then
                        MsgBox("Your cursor must be located at least one clear paragraph away from any existing tables, otherwise they'll merge in unexpected ways." + vbCrLf + vbCrLf + "Please relocate your insertion point and try again")
                    Else
                        frm = New frm_TableBuilder()
                        frm.Show()
                    End If
                    'dlg_InsertTable.Show()
                    'tbl = objGlobals.glb_get_wrdSelTbl()
                    'objPlHTable.Plh_Table_Convert_ToAATable(tbl)

                Case "grpTbls_setTableTextCustomColour", "grpTbls_fillCellsWithCustomColour"
                    strColourPickerMode = "text_Colour"
                    Select Case e.Control.Id
                        Case "grpTbls_setTableTextCustomColour"
                            strColourPickerMode = "text_Colour"
                        Case "grpTbls_fillCellsWithCustomColour"
                            strColourPickerMode = "tbl_Cells"
                    End Select
                    '
                    objGlobals.glb_show_ColorPicker(strColourPickerMode)
                    'frmPicker = New frm_colorPicker(strColourPickerMode)
                    'frmPicker.Show()

                Case "grpTbls_TableColumnHeadingsStyle", "grpTbls_TableColumnHeadingsStyle_small"
                    Call objStylesMgr.applyStyle_TableColumnHeadings()
                Case "grpTbls_TableUnitsRowStyle", "grpTbls_TableUnitsRowStyle_small"
                    Call objStylesMgr.applyStyle_TableUnitsRow()
                Case "grpTbls_TableTextStyle"
                    Call objStylesMgr.applyStyle_TableText()
                Case "grpTbls_TableTextStyle_small"
                    Call objStylesMgr.applyStyle_TableText_small()
                '
                Case "grpTbls_TableListBullet"
                    Call objStylesMgr.applyStyle_TableListBullet()
                Case "grpTbls_TableListBullet_small"
                    Call objStylesMgr.applyStyle_TableListBullet_small()
                Case "grpTbls_TableListBullet2"
                    Call objStylesMgr.applyStyle_TableListBullet2()
                Case "grpTbls_TableListBullet2_small"
                    Call objStylesMgr.applyStyle_TableListBullet2_small()
                Case "grpTbls_TableListBullet3"
                    Call objStylesMgr.applyStyle_TableListBullet3()
                Case "grpTbls_TableListBullet3_small"
                    Call objStylesMgr.applyStyle_TableListBullet3_small()
                '
                Case "grpTbls_ListNumber"
                    Call objStylesMgr.applyStyle_TableListNumber()
                Case "grpTbls_ListNumber_small"
                    Call objStylesMgr.applyStyle_TableListNumber_small()
                Case "grpTbls_ListNumber2"
                    Call objStylesMgr.applyStyle_TableListNumber2()
                Case "grpTbls_ListNumber2_small"
                    Call objStylesMgr.applyStyle_TableListNumber2_small()
                Case "grpTbls_ListNumber3"
                    Call objStylesMgr.applyStyle_TableListNumber3()
                Case "grpTbls_ListNumber3_small"
                    Call objStylesMgr.applyStyle_TableListNumber3_small()

                Case "grpTbls_Quote"
                    Call objStylesMgr.applyStyle_TableQuote()
                Case "grpTbls_QuoteListBullet"
                    Call objStylesMgr.applyStyle_TableQuoteBullet()
                Case "grpTbls_QuoteSource"
                    Call objStylesMgr.applyStyle_TableQuoteSource()
                'Call objStylesMgr.applyStyle_TableListNumber3_small()

                Case "grpTbls_Quote_small"
                    Call objStylesMgr.applyStyle_TableQuote_small()
                Case "grpTbls_QuoteListBullet_small"
                    Call objStylesMgr.applyStyle_TableQuoteBullet_small()
                Case "grpTbls_QuoteSource_small"
                    Call objStylesMgr.applyStyle_TableQuoteSource_small()

                Case "grpTbls_TableSideHeading1"
                    Call objStylesMgr.applyStyle_TableSideHeading1()
                Case "grpTbls_TableSideHeading1_small"
                    Call objStylesMgr.applyStyle_TableSideHeading1_small()
                Case "grpTbls_TableSideHeading2"
                    Call objStylesMgr.applyStyle_TableSideHeading2()
                Case "grpTbls_TableSideHeading2_small"
                    Call objStylesMgr.applyStyle_TableSideHeading2_small()
                '
                Case "grpTbls_StyleSet_TableQuote"
                    rng = Globals.ThisAddIn.Application.Selection.Range
                    objStylesMgr.insert_ExampleTableText_column01(rng, "normal")
                Case "grpTbls_StyleSet_TableListBullets"
                    rng = Globals.ThisAddIn.Application.Selection.Range
                    objStylesMgr.insert_ExampleTableText_column02(rng, "normal")
                Case "grpTbls_StyleSet_TableListNumbers"
                    rng = Globals.ThisAddIn.Application.Selection.Range
                    objStylesMgr.insert_ExampleTableText_column03(rng, "normal")
                '
                Case "grpTbls_StyleSet_TableQuote_small"
                    rng = Globals.ThisAddIn.Application.Selection.Range
                    objStylesMgr.insert_ExampleTableText_column01(rng, "small")
                Case "grpTbls_StyleSet_TableListBullets_small"
                    rng = Globals.ThisAddIn.Application.Selection.Range
                    objStylesMgr.insert_ExampleTableText_column02(rng, "small")
                Case "grpTbls_StyleSet_TableListNumbers_small"
                    rng = Globals.ThisAddIn.Application.Selection.Range
                    objStylesMgr.insert_ExampleTableText_column03(rng, "small")
                '
                Case "grpTbls_ColourCells"
                    '
                    If objGlobals.glb_get_wrdSel.Range.Tables.Count = 0 Then
                        MsgBox("Make certain that you have selected the cells that you wish to fill")
                        Exit Sub
                    End If
                    '
                    objTables.tbl_colour_set_colourOfCells(objGlobals.glb_get_wrdSel.Cells, objTables._glb_colour_UnitsGrey)

                Case "grpTbls_ColourHeadingsRow"
                    '
                    If objGlobals.glb_get_wrdSelRngAll.Tables.Count = 0 Then
                        Call objMsgMgr.colourRowsErrorMessage()
                        Exit Sub
                    End If
                    '
                    rng = objGlobals.glb_get_wrdSelRngAll
                    If rng.Rows.Count = 0 Then
                        Call objMsgMgr.colourRowsErrorMessage()
                        Exit Sub
                    End If
                    '
                    For Each dr In rng.Rows
                        Call objTables.tbl_colour_set_colourOfRow(dr, objTables._glb_colour_purple_Dark)
                    Next dr
            '
                Case "grpTbls_ColourUnitsRow"
                    If objGlobals.glb_get_wrdSelRngAll.Tables.Count = 0 Then
                        Call objMsgMgr.colourRowsErrorMessage()
                        Exit Sub
                    End If
                    '
                    rng = Globals.ThisAddIn.Application.Selection.Range
                    If rng.Rows.Count = 0 Then
                        Call objMsgMgr.colourRowsErrorMessage()
                        Exit Sub
                    End If
                    For Each dr In rng.Rows
                        Call objTables.tbl_colour_set_colourOfRow(dr, objTables._glb_colour_UnitsGrey)
                    Next dr
                Case "grpTbls_AllBorders"
                    rng = objGlobals.glb_get_wrdSelRng
                    '
                    objTables.tbl_borders_colourAndVisibility(rng, True, objTables._glb_colour_TableBorders)
            '
                Case "grpTbls_AllBordersRemove"
                    rng = objGlobals.glb_get_wrdSelRng
                    '
                    For Each tbl In objGlobals.glb_get_wrdSelRngAll.Tables
                        Call objTables.tbl_doBorders_MaintainPadding(tbl, False, objTables._glb_colour_TableBorders)
                    Next tbl
                    '
                Case "grpTbls_ApplyBottomBorder"
                    MsgBox("Still necessary, given the current table facilities..?")

                Case "grpTbls_convertTabletoES"
                    objPlHTable.PlhTbl_Captions_ConvertCaptionsTo_ES(objGlobals.glb_get_wrdSelRngAll)
                    objFlds.updateSequenceNumbers_Tables_ES()

                Case "grpTbls_convertTabletoStd"
                    objPlHTable.PlhTbl_Captions_ConvertCaptionsTo_Report(objGlobals.glb_get_wrdSelRngAll)
                    objFlds.updateSequenceNumbers_Tables()

                Case "grpTbls_convertTabletoApp"
                    objPlHTable.PlhTbl_Captions_ConvertCaptionsTo_Appendix(objGlobals.glb_get_wrdSelRngAll)
                    objFlds.updateSequenceNumbers_Tables_AP()

        '*******
                Case "grpTbls_convertTabletoLT"
                    objPlHTable.PlhTbl_Captions_ConvertCaptionsTo_Letter(objGlobals.glb_get_wrdSelRngAll)
                    objFlds.updateSequenceNumbers_Tables_LT()

                Case "grpTbls_convertTabletoX"
                Case Else
            End Select

        Catch ex As Exception

        End Try
        '
        objGlobals.glb_get_wrdApp.ScreenUpdating = True
        '


    End Sub
    '
    Private Sub PIF_Plh_grpTblsEdit_Click(sender As Object, e As RibbonControlEventArgs) Handles grpTblsEdit_InsertColumnRight.Click, grpTblsEdit_InsertColumnLeft.Click, grpTblsEdit_Delete_Column.Click, grpTblsEdit_InsertRowBelow.Click, grpTblsEdit_InsertRowAbove.Click, grpTblsEdit_Delete_Row.Click, grpTblsEdit_UndoTableAction.Click, grpTblsEdit_PastePriorTable.Click, grpTblsEdit_CopyTable.Click, grpTblsEdit_Split_Table.Click, grpTblsEdit_Convert_StdToEncaps.Click, grpTblsEdit_Convert_EncapsToStd.Click
        Dim objGlobals As New cGlobals()
        Dim objTblsMgr As New cTablesMgr()
        Dim objMsgMgr As New cMessageManager()
        Dim tbl, tblTop, tblBody, tblBottom As Word.Table
        Dim drCol As Word.Column
        Dim drCell, drCellSelected As Word.Cell
        Dim rng As Word.Range
        Dim splitParaTop, splitParaBottom As Word.Paragraph
        Dim strMsg, strMsg2 As String
        Dim columnIndex As Integer
        Dim tblPreferredWidth As Single
        '
        '
        columnIndex = -1
        strMsg = ""
        strMsg2 = "Error..."
        '
        drCol = Nothing
        tblTop = Nothing
        tblBody = Nothing
        tblBottom = Nothing
        splitParaTop = Nothing
        splitParaBottom = Nothing
        tblPreferredWidth = 100
        '
        objGlobals.glb_screen_update(False)
        '
        Try
            Select Case e.Control.Id
                Case "grpTblsEdit_InsertRowAbove"
                    strMsg = objTblsMgr.IsOKToAddTableColumn()
                    '
                    If strMsg = "" Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        '
                        If Not IsNothing(tbl) Then
                            objGlobals.glb_get_wrdSel.InsertRowsAbove(1)
                        Else
                            MsgBox("Please make certain that your cursor is in the row of interest in a Table")
                        End If
                    Else
                        MsgBox(strMsg,, "Warning")
                    End If
                    '
                Case "grpTblsEdit_InsertRowBelow"
                    strMsg = objTblsMgr.IsOKToAddTableColumn()
                    '
                    If strMsg = "" Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        '
                        If Not IsNothing(tbl) Then
                            objGlobals.glb_get_wrdSel.InsertRowsBelow(1)
                        Else
                            MsgBox("Please make certain that your cursor is in the row of interest in a Table")
                        End If
                    Else
                        MsgBox(strMsg,, "Warning")
                    End If
                    '
                Case "grpTblsEdit_InsertColumnRight"
                    strMsg = objTblsMgr.IsOKToAddTableColumn()
                    '
                    If strMsg = "" Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        '
                        If Not IsNothing(tbl) Then
                            drCellSelected = objGlobals.glb_get_wrdSelCell()
                            If Not IsNothing(drCellSelected) Then
                                strMsg = objTblsMgr.tbl_Encaps_SelectedIsFirstOrLastCell(drCellSelected, tbl)
                                '
                                If strMsg = "" Then
                                    rng = tbl.Range
                                    rng.Copy()
                                    columnIndex = objTblsMgr.tbl_aacTable_ColumnInsertDelete("right", tbl)

                                    If columnIndex < 0 Then
                                        MsgBox(strMsg2 + "columnIndex = " + columnIndex.ToString())
                                    Else
                                        '
                                        If columnIndex > 1 Then
                                            drCell = objTblsMgr.tbl_get_firstCellWithColumnIndex(columnIndex, tbl)
                                        Else
                                            'ColumnIndex is 1. In oreder to bypass the Caption row/cell which also has
                                            'a columnindex of 1 we need to go to the second cell
                                            drCell = objTblsMgr.tbl_get_secondCellWithColumnIndex(columnIndex, tbl)
                                        End If
                                        rng = drCell.Range
                                        rng.Collapse(WdCollapseDirection.wdCollapseStart)
                                        rng.Select()
                                    End If
                                    '
                                Else
                                    MsgBox(strMsg,, "Warning")
                                End If
                            Else
                                MsgBox("No table cell selected.." + vbCrLf + "Try moving your cursor into the body of the table.")
                            End If


                        Else
                            MsgBox("Your cursor must be in the Table that you want to edit")
                        End If
                        '
                    Else
                        MsgBox(strMsg)
                    End If
                '
                Case "grpTblsEdit_InsertColumnLeft"
                    strMsg = objTblsMgr.IsOKToAddTableColumn()
                    '
                    If strMsg = "" Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        '
                        If Not IsNothing(tbl) Then
                            drCellSelected = objGlobals.glb_get_wrdSelCell()
                            If Not IsNothing(drCellSelected) Then
                                strMsg = objTblsMgr.tbl_Encaps_SelectedIsFirstOrLastCell(drCellSelected, tbl)
                                '
                                If strMsg = "" Then
                                    rng = tbl.Range
                                    rng.Copy()
                                    columnIndex = objTblsMgr.tbl_aacTable_ColumnInsertDelete("left", tbl)

                                    If columnIndex < 0 Then
                                        MsgBox(MsgBox(strMsg2 + "columnIndex = " + columnIndex.ToString()))
                                    Else
                                        '
                                        If columnIndex > 1 Then
                                            drCell = objTblsMgr.tbl_get_firstCellWithColumnIndex(columnIndex, tbl)
                                        Else
                                            'ColumnIndex is 1. In oreder to bypass the Caption row/cell which also has
                                            'a columnindex of 1 we need to go to the second cell
                                            drCell = objTblsMgr.tbl_get_secondCellWithColumnIndex(columnIndex, tbl)
                                        End If
                                        rng = drCell.Range
                                        rng.Collapse(WdCollapseDirection.wdCollapseStart)
                                        rng.Select()
                                    End If
                                    '
                                Else
                                    MsgBox(strMsg,, "Warning")
                                End If
                            Else
                                MsgBox("No table cell selected.." + vbCrLf + "Try moving your cursor into the body of the table.")
                            End If


                        Else
                            MsgBox("Your cursor must be in the Table that you want to edit")
                        End If
                        '
                    Else
                        MsgBox(strMsg)
                    End If
                    '
                Case "grpTblsEdit_PastePriorTable"
                    Try
                        strMsg = objTblsMgr.IsOKToAddTableColumn()
                        '
                        If strMsg = "" Then
                            If System.Windows.Forms.Clipboard.ContainsText Then
                                rng = objGlobals.glb_get_wrdApp().Selection.Range
                                rng.Paste()
                            Else
                                MsgBox("The Clipboard does not contain a table.",, "Warning")
                            End If
                        Else
                            MsgBox(strMsg,, "Warning")
                        End If
                        '
                    Catch ex As Exception
                        MsgBox("The Paste failed...")
                    End Try
                    '
                Case "grpTblsEdit_UndoColumnInsert", "grpTblsEdit_UndoTableAction"
                    strMsg = objTblsMgr.tbl_isOK_ToPasteOverTable()
                    '
                    If strMsg = "" Then
                        If objMsgMgr.doTableColumnUndoMessage() Then
                            rng = objGlobals.glb_get_wrdApp.Selection.Range
                            '
                            Try
                                If System.Windows.Forms.Clipboard.ContainsText Then
                                    If rng.Tables.Count <> 0 Then
                                        If System.Windows.Forms.Clipboard.ContainsText Then
                                            tbl = rng.Tables.Item(1)
                                            tbl.Delete()
                                            rng = objGlobals.glb_get_wrdApp.Selection.Range
                                            rng.Paste()
                                        Else
                                            MsgBox("The Undo failed... The ClipBoard was empty")
                                        End If
                                        '
                                    Else
                                        MsgBox("The Undo failed... Your cursor was Not In a table (i.e. the table you want To replace")

                                        'rng = Globals.ThisDocument.Application.Selection.Range
                                        'rng.Paste()
                                    End If
                                Else
                                    MsgBox("The ClipBoard was empty, so no action was taken")
                                End If
                            Catch ex As Exception
                                MsgBox("The Undo failed... Probably because the Table has been cleared from the ClipBoard")
                            End Try
                        End If

                    Else
                        MsgBox(strMsg)
                    End If
                    '
                '
                Case "grpTblsEdit_Delete_Column"
                    'objGlobals.glb_screen_update(True)
                    '
                    'Note that when we delete a column, the columnIndex may point to a column that does not exist
                    '
                    strMsg = objTblsMgr.IsOKToAddTableColumn()
                    '
                    If strMsg = "" Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        '
                        If Not IsNothing(tbl) Then
                            drCellSelected = objGlobals.glb_get_wrdSelCell()
                            If Not IsNothing(drCellSelected) Then
                                strMsg = objTblsMgr.tbl_Encaps_SelectedIsFirstOrLastCell(drCellSelected, tbl)
                                '
                                If strMsg = "" Then
                                    rng = tbl.Range
                                    rng.Copy()
                                    'In this case columnIndex is set to the column that was deleted.. This is not problem unless
                                    'it was the last column
                                    Try
                                        'columnIndex = objTblsMgr.tbl_aacTable_ColumnDelete(tbl)
                                        columnIndex = objTblsMgr.tbl_aacTable_ColumnInsertDelete("delete", tbl)
                                        drCell = objTblsMgr.tbl_get_firstCellWithColumnIndex(columnIndex, tbl)
                                        '
                                        If Not IsNothing(drCell) Then
                                            rng = drCell.Range
                                            rng.Collapse(WdCollapseDirection.wdCollapseStart)
                                            rng.Select()
                                        End If

                                        'If IsNothing(drCell) Then
                                        'The column does not exist
                                        'columnIndex = columnIndex - 1
                                        'Else
                                        'drCell.Range.Delete()
                                        'End If
                                    Catch ex As Exception

                                    End Try
                                    '
                                Else
                                    MsgBox(strMsg,, "Warning")
                                End If
                            Else
                                MsgBox("No table cell selected.." + vbCrLf + "Try moving your cursor into the body Of the table.")
                            End If
                        Else
                            MsgBox("Your cursor must be In the Table that you want To edit")
                        End If
                        '
                    Else
                        MsgBox(strMsg)
                    End If
                    '
                Case "grpTblsEdit_Delete_Row"
                    strMsg = objTblsMgr.IsOKToAddTableColumn()
                    '
                    If strMsg = "" Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        '
                        If Not IsNothing(tbl) Then
                            objGlobals.glb_get_wrdSel.Rows.Delete()
                        Else
                            MsgBox("Your cursor must be In the Table that you want To edit")
                        End If
                    Else
                        MsgBox(strMsg,, "Warning")
                    End If
                    '

                    '
                Case "grpTblsEdit_Convert_EncapsToStd"
                    strMsg = objTblsMgr.IsOKToAddTableColumn()
                    '
                    Try
                        If strMsg = "" Then
                            tbl = objGlobals.glb_get_wrdSelTbl()
                            If Not IsNothing(tbl) Then
                                If objTblsMgr.tbl_is_EncapsulatedTable(tbl) Or objTblsMgr.tbl_is_EncapsulatedFigure(tbl) Or objTblsMgr.tbl_is_EncapsulatedBox(tbl) Then
                                    'If objTblsMgr.glb_tbls_isRegularByRow(tbl) Then
                                    tblBody = objTblsMgr.tbl_convert_tblEncapsToStd(tbl)
                                    drCell = tbl.Range.Cells.Item(1)
                                    rng = drCell.Range
                                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                                    rng.Select()
                                    'Else
                                    'MsgBox("The Selected table must be regular by row (i.e. no merged column cells).")
                                    'End If
                                Else
                                    MsgBox("Please make certain that your cursor Is In a 'Encapsulated' table (i.e. a table with integrated 'Caption' and 'Source/Note' rows.")

                                End If
                            Else
                                MsgBox("Please make certain that your cursor is in the table you wish to convert.")

                            End If
                            '
                        Else
                            MsgBox(strMsg)
                        End If
                    Catch ex As Exception

                    End Try
                    '
                Case "grpTblsEdit_Convert_StdToEncaps"
                    '
                    strMsg = objTblsMgr.IsOKToAddTableColumn()
                    '
                    Try
                        If strMsg = "" Then
                            tbl = objGlobals.glb_get_wrdSelTbl()
                            '
                            If Not IsNothing(tbl) Then
                                If Not (objTblsMgr.tbl_is_EncapsulatedTable(tbl) Or objTblsMgr.tbl_is_EncapsulatedFigure(tbl) Or objTblsMgr.tbl_is_EncapsulatedBox(tbl)) Then
                                    If objTblsMgr.tbl_is_tblStandard(tbl) Then
                                        '
                                        If objTblsMgr.tbl_has_topCaption(tbl) Then
                                            If objTblsMgr.tbl_has_SourceNotePara(tbl) Then
                                                strMsg = objTblsMgr.tbl_convert_tblStdToEncaps(tbl)
                                                drCell = tbl.Range.Cells.Item(2)
                                                rng = drCell.Range
                                                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                                                rng.Select()
                                            Else
                                                If objMsgMgr.msg_stdTbl_hasNoSourceNoteRow() Then
                                                    strMsg = objTblsMgr.tbl_convert_tblStdToEncaps(tbl)
                                                    drCell = tbl.Range.Cells.Item(2)
                                                    rng = drCell.Range
                                                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                                                    rng.Select()
                                                Else
                                                    MsgBox("The action has been cancelled")
                                                End If
                                            End If
                                        Else
                                            MsgBox("This action cannot proceed because the table has no caption")
                                        End If

                                    End If
                                Else
                                    MsgBox("Please make certain that your cursor is in a 'Standard' table (i.e. a table with separate Caption and Source/Note rows.")

                                End If

                            Else
                                MsgBox("Please make certain that your cursor is in the body of the table.")
                            End If

                        Else
                            MsgBox(strMsg)
                        End If

                    Catch ex As Exception

                    End Try
                    '


                Case "grpTblsEdit_Split_Table"
                    '
                    strMsg = objTblsMgr.IsOKToAddTableColumn()
                    '
                    If strMsg = "" Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        If Not IsNothing(tbl) Then
                            'This will handle eny structure in the body
                            '
                            If objTblsMgr.tbl_is_EncapsulatedTable(tbl) Or objTblsMgr.tbl_is_EncapsulatedFigure(tbl) Then
                                tblBody = objTblsMgr.tbl_Encaps_split(tbl, tblTop, tblBottom, splitParaTop, splitParaBottom)
                                'tblBody = objTblsMgr.tbl_Encaps_splitTop(tbl, tblTop, splitParaTop)
                                'tblBody = objTblsMgr.tbl_Encaps_splitBottom(tbl, tblBottom, splitParaBottom)
                                'objTblsMgr.tbl_colour_set_colourOfCellsToNone(tblBottom.Range.Cells)
                                '
                                'rng = splitParaBottom.Range
                                'rng.Select()
                                '
                                'rng = tblBottom.Range
                                'rng.Select()
                                '
                                'If objTblsMgr.glb_tbls_isRegularByRow(tbl) Then
                                'tblBody = objTblsMgr.tbl_Encaps_split(tbl, tblTop, tblBottom, splitParaTop, splitParaBottom)
                                'objTblsMgr.tbl_colour_set_colourOfCellsToNone(tblBottom.Range.Cells)
                                'Else
                                'MsgBox("The Selected table must be regular by row (i.e. no merged column cells).")
                                'End If
                            Else
                                MsgBox("Please make certain that your cursor is in an 'Encapsulated Table.'")
                            End If
                        Else
                            MsgBox("Please make certain that your cursor is in the table you wish to split")
                        End If

                    Else
                        MsgBox(strMsg)
                    End If
                    '
                Case "grpTblsEdit_StdToEncaps_BuildTopRow"
                    '
                    If objGlobals.glb_get_wrdSel().Tables.Count >= 1 Then
                        tbl = objGlobals.glb_get_wrdSel().Tables.Item(1)
                        objTblsMgr.tbl_build_topRowAsCells(tbl)

                    End If
                    '
                Case "grpTblsEdit_CopyTable"
                    If objGlobals.glb_selection_IsInTable() Then
                        tbl = objGlobals.glb_get_wrdSelTbl2()
                        rng = tbl.Range
                        rng.Copy()
                        '
                        MsgBox("The table has been sucessfully copied")
                    Else
                        MsgBox("Please place your cursor in the table you wish to copy to the clipboard")
                    End If

                Case Else
            End Select

        Catch ex As Exception

        End Try

        objGlobals.glb_screen_update(True)
        '
    End Sub
    '
    Private Sub PIF_Plh_grpTblsNote_Click(sender As Object, e As RibbonControlEventArgs) Handles grpTblsPlh_SourceLabelAndStyle.Click, grpTblsPlh_SourceForOverType.Click, grpTblsPlh_NoteLabelAndStyle.Click
        Dim objGlobals As New cGlobals()
        Dim objTblsMgr As New cTablesMgr()
        Dim objMsgMgr As New cMessageManager()
        Dim rng As Word.Range
        Dim strMsg As String
        '
        objGlobals.glb_screen_update(False)
        strMsg = ""
        '
        Select Case e.Control.Id
            Case "grpTblsPlh_SourceLabelAndStyle"
                rng = objTblsMgr.tbl_insert_SourceAndNoteText(objGlobals.glb_get_wrdSelRng, "sourceOnly")
                rng.Select()
            Case "grpTblsPlh_NoteLabelAndStyle"
                rng = objTblsMgr.tbl_insert_SourceAndNoteText(objGlobals.glb_get_wrdSelRng, "note")
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                rng.Select()
            Case "grpTblsPlh_SourceForOverType"
                rng = objTblsMgr.tbl_insert_SourceAndNoteText(objGlobals.glb_get_wrdSelRng, "sourceAndNote")
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                rng.Select()

            Case "grpTblsNte_NoteStyle"

            Case "grpTblsNte_NoteLabelStyle"

            Case "grpTblsPlh_NoteText"
                'rng = objTblsMgr.tbl_insert_SourceAndNoteText(objGlobals.glb_get_wrdSelRng, "note")
                'rng.Select()
                '
            Case Else
        End Select

        objGlobals.glb_screen_update(True)
        '
    End Sub

    Private Sub PIF_Plh_grpTblsPlh_Click(sender As Object, e As RibbonControlEventArgs) Handles grpTblsPlh_HeadingAndSource.Click, grpTblsPlh_HeadingAndSourceApp.Click, grpTblsPlh_HeadingAndSourceES.Click, grpTblsPlh_CaptionAndHeading.Click, grpTblsPlh_CaptionAndHeadingApp.Click, grpTblsPlh_CaptionAndHeadingES.Click, grpTblsPlh_AddTable_Simple.Click, grpTblsPlh_DeleteTable_fast.Click, grpTblsPlh_DeleteTable.Click, grpTbl_mnu_AAPlh_To_HalfPlh_Right.Click, grpTbl_mnu_AAPlh_To_HalfPlh_Left.Click, grpTbl_mnu_AAPlh_Reset_to_FullColumn.Click, grpTblsPlh_rapidFormat_Encapsulated.Click, grpTblsPlh_rapidFormat.Click, grpBoxes_mnu_rapidFormat_EncapTbl_Force_Body.Click, grpBoxes_mnu_rapidFormat_StdTbl_Force_ES.Click, grpBoxes_mnu_rapidFormat_StdTbl_Force_Body.Click, grpBoxes_mnu_rapidFormat_StdTbl_Force_AP.Click, grpBoxes_mnu_rapidFormat_EncapTbl_Force_LT.Click, grpBoxes_mnu_rapidFormat_EncapTbl_Force_ES.Click, grpBoxes_mnu_rapidFormat_EncapTbl_Force_AP.Click, grpBoxes_mnu_rapidFormat_StdTbl_Force_LT.Click
        Dim rng, oldRange As Range
        Dim objWrks As New cWorkArounds()
        Dim objPlHTbl As New cPlHTable()
        Dim objPlhBase As New cPlHBase()

        Dim objFlds As New cFieldsMgr()
        Dim objGlobals As New cGlobals()
        Dim objTools As New cTools()
        Dim objStylesMgr As New cStylesManager()
        Dim objTblStylesMgr As New cTableStyles()
        Dim objCaptMgr As New cCaptionManager()
        Dim objIsOK As New cIsOKToDo()
        Dim tblWidth As Single
        Dim tbl, hf_tbl As Word.Table
        Dim sect As Word.Section
        Dim myDoc As Word.Document
        Dim strMsg As String

        Dim objTblMgr As cTablesMgr

        'Dim dlg_InsertTable As Word.Dialog = Globals.ThisAddIn.Application.Dialogs(Word.WdWordDialog.wdDialogTableInsertTable)
        Dim dlg_InsertTable As Word.Dialog = objGlobals.glb_get_wrdApp.Dialogs(Word.WdWordDialog.wdDialogTableInsertTable)
        '
        objTools = New cTools
        strMsg = ""
        '
        objGlobals.glb_screen_update(False)
        '
        hf_tbl = Nothing
        sect = objGlobals.glb_get_wrdSect()
        myDoc = sect.Range.Document
        oldRange = objGlobals.glb_get_wrdSelRngAll()
        '
        tblWidth = objGlobals.glb_get_widthBetweenMargins(sect) / 2
        '
        Try
            objTblMgr = New cTablesMgr()
            Select Case e.Control.Id
                Case "grpTbl_mnu_HalfTables_Left"
                    tbl = objGlobals.glb_get_wrdSelTbl()
                    objTblMgr.tbl_format_rapidFormat_Encap(tbl)
                    '
                    objTblMgr.tbl_width_Change(tbl, tblWidth)
                    objTblMgr.tbl_convert_toFloating(tbl)
                    tbl.Rows.WrapAroundText = True
                    tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
                    tbl.Rows.Alignment = WdRowAlignment.wdAlignRowLeft

                Case "grpTbl_mnu_AAPlh_To_HalfPlh_Left"
                    If objIsOK.isOKto_doAction_inReportBody() = objIsOK._isOK Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        If Not IsNothing(tbl) Then
                            objTblMgr.tbl_convert_toFloatingRelToMargin(tbl, "left")
                            'The following is done inside the tbl_convert method above
                            'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                            'tbl.PreferredWidth = 50
                        Else
                            MsgBox("Please make certain your cursor is in a pre built PlaceHolder")
                        End If
                    Else
                        MsgBox("This action is only supported on AA PlaceHolders in the body of the report")
                    End If
                    '
                Case "grpTbl_mnu_AAPlh_Reset_to_FullColumn"
                    If objIsOK.isOKto_doAction_inReportBody() = objIsOK._isOK Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        If Not IsNothing(tbl) Then
                            objTblMgr.tbl_convert_toInLine(tbl)
                            tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                            tbl.PreferredWidth = 100
                            'tbl.Rows.Alignment = WdRowAlignment.wdAlignRowRight
                            '
                        Else
                            MsgBox("Please make certain your cursor is in a pre built PlaceHolder")
                        End If
                    Else
                        MsgBox("This action is only supported on AA PlaceHolders in the body of the report")
                    End If

                Case "grpTbl_mnu_AAPlh_To_HalfPlh_Right"
                    If objIsOK.isOKto_doAction_inReportBody() = objIsOK._isOK Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        If Not IsNothing(tbl) Then
                            objTblMgr.tbl_convert_toFloatingRelToMargin(tbl, "right")
                            'tbl.Rows.Alignment = WdRowAlignment.wdAlignRowRight
                            '
                        Else
                            MsgBox("Please make certain your cursor is in a pre built PlaceHolder")
                        End If
                    Else
                        MsgBox("This action is only supported on AA PlaceHolders in the body of the report")
                    End If
                    '
                Case "grpTbl_mnu_HalfTables_Right"
                    tbl = objGlobals.glb_get_wrdSelTbl()
                    objTblMgr.tbl_format_rapidFormat_Encap(tbl)
                    '
                    objTblMgr.tbl_width_Change(tbl, tblWidth)
                    objTblMgr.tbl_convert_toFloating(tbl)
                    tbl.Rows.WrapAroundText = True
                    tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
                    tbl.Rows.Alignment = WdRowAlignment.wdAlignRowRight
                    '

                Case "grpTblsPlh_rapidFormat"
                    If objIsOK.isOKto_doAction_inReportBody() = objIsOK._isOK Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        'tbl = objGlobals.glb_get_wrdSelTbl2()
                        If Not IsNothing(tbl) Then
                            'objTblMgr.tbl_format_rapidFormat2(tbl)
                            objTblMgr.tbl_format_rapidFormat(tbl)
                            tbl.Rows.First.HeadingFormat = True             'Causes Row 1 to repeat
                            'tbl.Rows.Item(1).HeadingFormat = True
                            'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                            'tbl.PreferredWidth = 100
                        Else
                            MsgBox("Please make certain your cursor is in a pre built (Ribbon Insert > Table) unformatted table")
                        End If
                    Else
                        MsgBox("This action is only supported on pre built tables in the body of the report")
                    End If
                    'tbl.AllowPageBreaks = True
                    'tbl.Rows.AllowBreakAcrossPages = False
                    'objTblMgr.tbl_fix_Table(tbl, True, RGB(255, 0, 0), True)
                    '
                Case "grpTblsPlh_rapidFormat_Encapsulated"
                    If objIsOK.isOKto_doAction_inReportBody() = objIsOK._isOK Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        'tbl = objGlobals.glb_get_wrdSelTbl2()
                        If Not IsNothing(tbl) Then
                            rng = objTblMgr.tbl_format_rapidFormat_Encap(tbl)
                            'The table returned is not the full table, but the table sandwiched
                            'between the upper and lower cell
                            tbl = rng.Tables.Item(1)
                            tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                            tbl.PreferredWidth = 100
                        Else
                            MsgBox("Please make certain your cursor is in a pre built (Ribbon Insert > Table) unformatted table")
                        End If
                    Else
                        MsgBox("This action is only supported on pre built tables in the body of the report")
                    End If
                    '
                Case "grpBoxes_mnu_rapidFormat_StdTbl_Force_LT"
                    strMsg = objIsOK.isOKto_doAction_inReportBody()
                    If strMsg = objIsOK._isOK Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        'tbl = objGlobals.glb_get_wrdSelTbl2()
                        If Not IsNothing(tbl) Then
                            rng = objTblMgr.tbl_format_rapidFormat(tbl, "LT")
                            tbl.Rows.First.HeadingFormat = True             'Causes Row 1 to repeat
                            'The table returned is not the full table,
                            'but the table sandwiched
                            'between the upper and lower cell
                            'tbl = rng.Tables.Item(1)
                            'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                            'tbl.PreferredWidth = 100
                        Else
                            MsgBox("Please make certain your cursor is in a pre built (Ribbon Insert > Table) unformatted table")
                        End If
                    Else
                        MsgBox(strMsg)
                    End If
                Case "grpBoxes_mnu_rapidFormat_StdTbl_Force_ES"
                    strMsg = objIsOK.isOKto_doAction_inReportBody()
                    If strMsg = objIsOK._isOK Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        'tbl = objGlobals.glb_get_wrdSelTbl2()
                        If Not IsNothing(tbl) Then
                            rng = objTblMgr.tbl_format_rapidFormat(tbl, "ES")
                            tbl.Rows.First.HeadingFormat = True                             'Causes Row 1 to repeat
                            'The table returned is not the full table, but the
                            'table sandwiched between the upper and lower cell
                            'tbl = rng.Tables.Item(1)
                            'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                            'tbl.PreferredWidth = 100
                        Else
                            MsgBox("Please make certain your cursor is in a pre built (Ribbon Insert > Table) unformatted table")
                        End If
                    Else
                        MsgBox(strMsg)
                    End If


                Case "grpBoxes_mnu_rapidFormat_StdTbl_Force_Body"
                    strMsg = objIsOK.isOKto_doAction_inReportBody()
                    If strMsg = objIsOK._isOK Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        'tbl = objGlobals.glb_get_wrdSelTbl2()
                        If Not IsNothing(tbl) Then
                            rng = objTblMgr.tbl_format_rapidFormat(tbl, "BD")
                            tbl.Rows.First.HeadingFormat = True                             'Causes Row 1 to repeat
                            'The table returned is not the full table, but the
                            'table sandwiched between the upper and lower cell
                            'tbl = rng.Tables.Item(1)
                            'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                            'tbl.PreferredWidth = 100
                        Else
                            MsgBox("Please make certain your cursor is in a pre built (Ribbon Insert > Table) unformatted table")
                        End If
                    Else
                        MsgBox(strMsg)
                    End If

                Case "grpBoxes_mnu_rapidFormat_StdTbl_Force_AP"
                    strMsg = objIsOK.isOKto_doAction_inReportBody()
                    If strMsg = objIsOK._isOK Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        'tbl = objGlobals.glb_get_wrdSelTbl2()
                        If Not IsNothing(tbl) Then
                            rng = objTblMgr.tbl_format_rapidFormat(tbl, "AP")
                            tbl.Rows.First.HeadingFormat = True                             'Causes Row 1 to repeat
                            'The table returned is not the full table, but the
                            'table sandwiched between the upper and lower cell
                            'tbl = rng.Tables.Item(1)
                            'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                            'tbl.PreferredWidth = 100
                        Else
                            MsgBox("Please make certain your cursor is in a pre built (Ribbon Insert > Table) unformatted table")
                        End If
                    Else
                        MsgBox(strMsg)
                    End If
                    '
                Case "grpBoxes_mnu_rapidFormat_EncapTbl_Force_LT"
                    strMsg = objIsOK.isOKto_doAction_inReportBody()
                    If strMsg = objIsOK._isOK Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        'tbl = objGlobals.glb_get_wrdSelTbl2()
                        If Not IsNothing(tbl) Then
                            rng = objTblMgr.tbl_format_rapidFormat_Encap(tbl, "LT")
                            'The table returned is not the full table, but the table sandwiched
                            'between the upper and lower cell
                            'tbl = rng.Tables.Item(1)
                            'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                            'tbl.PreferredWidth = 100
                        Else
                            MsgBox("Please make certain your cursor is in a pre built (Ribbon Insert > Table) unformatted table")
                        End If
                    Else
                        MsgBox(strMsg)
                    End If
                Case "grpBoxes_mnu_rapidFormat_EncapTbl_Force_ES"
                    strMsg = objIsOK.isOKto_doAction_inReportBody()
                    If strMsg = objIsOK._isOK Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        'tbl = objGlobals.glb_get_wrdSelTbl2()
                        If Not IsNothing(tbl) Then
                            rng = objTblMgr.tbl_format_rapidFormat_Encap(tbl, "ES")
                            'The table returned is not the full table, but the table sandwiched
                            'between the upper and lower cell
                            'tbl = rng.Tables.Item(1)
                            'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                            'tbl.PreferredWidth = 100
                        Else
                            MsgBox("Please make certain your cursor is in a pre built (Ribbon Insert > Table) unformatted table")
                        End If
                    Else
                        MsgBox(strMsg)
                    End If


                Case "grpBoxes_mnu_rapidFormat_EncapTbl_Force_Body"
                    strMsg = objIsOK.isOKto_doAction_inReportBody()
                    If strMsg = objIsOK._isOK Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        'tbl = objGlobals.glb_get_wrdSelTbl2()
                        If Not IsNothing(tbl) Then
                            rng = objTblMgr.tbl_format_rapidFormat_Encap(tbl, "BD")
                            'The table returned is not the full table, but the table sandwiched
                            'between the upper and lower cell
                            'tbl = rng.Tables.Item(1)
                            'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                            'tbl.PreferredWidth = 100
                        Else
                            MsgBox("Please make certain your cursor is in a pre built (Ribbon Insert > Table) unformatted table")
                        End If
                    Else
                        MsgBox(strMsg)
                    End If

                Case "grpBoxes_mnu_rapidFormat_EncapTbl_Force_AP"
                    strMsg = objIsOK.isOKto_doAction_inReportBody()
                    If strMsg = objIsOK._isOK Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        'tbl = objGlobals.glb_get_wrdSelTbl2()
                        If Not IsNothing(tbl) Then
                            rng = objTblMgr.tbl_format_rapidFormat_Encap(tbl, "AP")
                            'The table returned is not the full table, but the table sandwiched
                            'between the upper and lower cell
                            'tbl = rng.Tables.Item(1)
                            'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                            'tbl.PreferredWidth = 100
                        Else
                            MsgBox("Please make certain your cursor is in a pre built (Ribbon Insert > Table) unformatted table")
                        End If
                    Else
                        MsgBox(strMsg)
                    End If


                Case "grpTblsPlh_ConvertAllTablesInDoc", "grpTestTblsPlh_ConvertAllTablesInDoc"
                    rng = objGlobals.glb_get_wrdActiveDoc.Range
                    objTblMgr.tbl_convert_aacTablesNoOutDent(rng)
                    objGlobals.glb_screen_update(True)
                    '
                Case "grpTblsPlh_ConvertAllTablesInRange", "grpTestTblsPlh_ConvertAllTablesInRange"
                    rng = objGlobals.glb_get_wrdSelRngAll()
                    objTblMgr.tbl_convert_aacTablesNoOutDent(rng)
                    objGlobals.glb_screen_update(True)
                    '
                Case "grpTblsPlh_toggleTableWidth"
                    '
                    strMsg = objIsOK.isOKto_toggle_PlhWidth()
                    '
                    If strMsg = "ok" Then
                        tbl = objGlobals.glb_get_wrdSelTbl()
                        If IsNothing(tbl) Then
                            MsgBox("Please make certain that your cursor is in a table based placeholder")
                        Else
                            If Not objTblMgr.tbl_toggle_tblWidth(tbl) Then
                                MsgBox("Error while attempting to toggle the placeholder width")
                            Else
                                'objPlhBase.Plh_scale_FigureImageShape(tbl)
                            End If
                        End If
                    Else
                        MsgBox(strMsg)
                    End If
                    '
                    '
                    objGlobals.glb_screen_update(True)
                    '
                Case "grpTblsPlh_TableWidthToWide", "grpTestTblsPlh_TableWidthToWide"
                    tbl = objGlobals.glb_get_wrdSelRngAll().Tables.Item(1)
                    objTblMgr.tbl_setWidth_toWide(tbl)
                Case "grpTestTblsPlh_TableWidthToStd"
                    tbl = objGlobals.glb_get_wrdSelRngAll().Tables.Item(1)
                    objTblMgr.tbl_setWidth_ToStandard(tbl)
                Case "grpTblsPlh_aacTblConvert", "grpTestTblsPlh_aacTblConvert"
                    '
                    '****
                    tbl = objGlobals.glb_get_wrdSelRngAll().Tables.Item(1)
                    'objTblMgr.tbl_set_ToWide(tbl)
                    objTblMgr.tbl_convert_aacToNoOutDent_ConvertOrResize(tbl)
                    'objTblMgr.tbl_convert_aacToNoOutDent(tbl)
                    objGlobals.glb_get_wrdApp.ScreenUpdating = True
                    objGlobals.glb_screen_update(True)

                Case "grpTblsPlh_HeadingAndSource"
                    '
                    objPlHTbl.PlhTbl_insert_Table("Table")
                    objFlds.upDateTableOfFigures()
                    '
                Case "grpTblsPlh_HeadingAndSource_17p5"
                    '                    '
                    objPlHTbl.PlhTbl_insert_TableWide("Table")
                    objFlds.upDateTableOfFigures()
                    '
                Case "grpTblsPlh_HeadingAndSourceApp"
                    '
                    objPlHTbl.PlhTbl_insert_Table("Table_AP")
                    objFlds.upDateTableOfFigures()
                    '
                Case "grpTblsPlh_HeadingAndSourceApp_17p5"
                    '                    '
                    objPlHTbl.PlhTbl_insert_TableWide("Table_AP")
                    objFlds.upDateTableOfFigures()
                    '
                Case "grpTblsPlh_HeadingAndSourceES"
                    '
                    objPlHTbl.PlhTbl_insert_Table("Table_ES")
                    objFlds.upDateTableOfFigures()
                    '
                Case "grpTblsPlh_HeadingAndSourceES_17p5"
                    '                    '
                    objPlHTbl.PlhTbl_insert_TableWide("Table_ES")
                    objFlds.upDateTableOfFigures()
                    '
                Case "grpTblsPlh_CaptionAndHeading"
                    rng = objPlHTbl.Plh_Captions_InsertCaptions("Table", objGlobals.glb_get_wrdSelRng, True)
                    objFlds.upDateTableOfFigures()
                    'rng = objChptPlhMgr.Plh_Captions_InsertCaptions("Table_AP", Globals.ThisDocument.Application.Selection.Range, True)
                    rng.Select()
                    '
                Case "grpTblsPlh_CaptionAndHeadingApp"
                    rng = objPlHTbl.Plh_Captions_InsertCaptions("Table_AP", objGlobals.glb_get_wrdSelRng, True)
                    objFlds.upDateTableOfFigures()
                    'rng = objChptPlhMgr.Plh_Captions_InsertCaptions("Table_AP", Globals.ThisDocument.Application.Selection.Range, True)
                    rng.Select()
                    '
                Case "grpTblsPlh_CaptionAndHeadingES"
                    rng = objPlHTbl.Plh_Captions_InsertCaptions("Table_ES", objGlobals.glb_get_wrdSelRng, True)
                    objFlds.upDateTableOfFigures()
                    'rng = objChptPlhMgr.Plh_Captions_InsertCaptions("Table_AP", Globals.ThisDocument.Application.Selection.Range, True)
                    rng.Select()
                    '
                Case "grpTblsPlh_DeleteTable_fast"
                    objTblMgr.tbl_delete_table()
            '
                Case "grpTblsPlh_DeleteTable"
                    objTblMgr.tbl_delete_table()
                    objFlds.upDateTableOfFigures()
                    '
                Case "grpTblsPlh_AddTable_Simple"
                    dlg_InsertTable.Show()
                    'Globals.ThisDocument.Application.Dialogs(Word.WdWordDialog.wdDialogTableInsertTable)

                    'Globals.ThisDocument.Application.Dialogs(Word.WdWordDialog.wdDialogTableInsertTable).Show()

                    Try
                        tbl = objGlobals.glb_get_wrdApp.Selection.Tables.Item(1)
                        tbl.Range.Style = objTblMgr.var_tbl_TextStyle

                        tbl.TopPadding = objTblMgr.var_tbl_padding_Top
                        tbl.BottomPadding = objTblMgr.var_tbl_padding_Bottom
                        tbl.LeftPadding = objTblMgr.var_tbl_padding_Left
                        tbl.RightPadding = objTblMgr.var_tbl_padding_Right
                        '
                        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
                        tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed)
                        '
                        'tblWidth = Me.adjustTable(tbl, objToolsMgr)
                        '
                        'Call Me.Table_doBorders_MaintainPadding(tbl, Me.chkBx_doBorders.Checked, Me.objGlobals.colour_TableBorders)

                    Catch ex As Exception

                    End Try
                Case Else
            End Select

        Catch ex1 As Exception

        End Try
        '
        objWrks.wrk_fix_forCursorRace()
        '
        objGlobals.glb_screen_update(True)
        '
        objWrks.wrk_fix_forCursorRace()
        '
        objGlobals.glb_screen_update(True)
        '
    End Sub

    Private Sub PIF_Plh_grpTbls_Click(sender As Object, e As RibbonControlEventArgs) Handles grpTbls_TableTextStyle.Click, grpTbls_TableUnitsRowStyle.Click, grpTbls_TableListBullet.Click, grpTbls_TableColumnHeadingsStyle.Click, grpTbls_TableListBullet3.Click, grpTbls_TableListBullet2.Click, grpTbls_ListNumber3.Click, grpTbls_ListNumber2.Click, grpTbls_ListNumber.Click, grpTbls_TableSideHeading1.Click, grpTbls_TableSideHeading2.Click, grpTbls_QuoteListBullet.Click, grpTbls_Quote.Click, grpTbls_QuoteSource.Click, grpTbls_StyleSet_TableQuote.Click, grpTbls_StyleSet_TableListBullets.Click, grpTbls_ColourUnitsRow.Click, grpTbls_ColourHeadingsRow.Click, grpTbls_ColourCells.Click, grpTbls_AllBordersRemove.Click, grpTbls_AllBorders.Click, grpTbls_convertTabletoStd.Click, grpTbls_convertTabletoLT.Click, grpTbls_convertTabletoES.Click, grpTbls_convertTabletoApp.Click, grpTbls_TableTextStyle_small.Click, grpTbls_TableUnitsRowStyle_small.Click, grpTbls_TableColumnHeadingsStyle_small.Click, grpTbls_ListNumber3_small.Click, grpTbls_ListNumber2_small.Click, grpTbls_ListNumber_small.Click, grpTbls_TableListBullet3_small.Click, grpTbls_TableListBullet2_small.Click, grpTbls_TableListBullet_small.Click, grpTbls_TableSideHeading2_small.Click, grpTbls_TableSideHeading1_small.Click, grpTbls_Quote_small.Click, grpTbls_QuoteSource_small.Click, grpTbls_QuoteListBullet_small.Click, grpTbls_StyleSet_TableQuote_small.Click, grpTbls_StyleSet_TableListBullets_small.Click, grpTbls_StyleSet_TableListNumbers_small.Click
        Dim frm As frm_TableBuilder
        Dim objGlobals As New cGlobals()
        Dim objStylesMgr As New cStylesManager
        Dim objPlHTable As New cPlHTable()
        Dim objTools As New cTools()
        Dim objFlds As New cFieldsMgr()
        Dim objMsgMgr As New cMessageManager()
        Dim strColourPickerMode As String
        'Dim frmPicker As frm_colorPicker
        Dim rng As Word.Range
        Dim objTables As New cTablesMgr()
        Dim tbl As Word.Table
        'Dim dlg_InsertTable As Word.Dialog = Globals.ThisDocument.Application.Dialogs(Word.WdWordDialog.wdDialogTableInsertTable)
        '
        objGlobals.glb_get_wrdApp.ScreenUpdating = False
        '
        Try
            Select Case e.Control.Id

                Case "grpTbls_setTableTextCustomColour", "grpTbls_fillCellsWithCustomColour"
                    strColourPickerMode = "text_Colour"
                    Select Case e.Control.Id
                        Case "grpTbls_setTableTextCustomColour"
                            strColourPickerMode = "text_Colour"
                        Case "grpTbls_fillCellsWithCustomColour"
                            strColourPickerMode = "tbl_Cells"
                    End Select
                    objGlobals.glb_show_ColorPicker(strColourPickerMode)
                    'frmPicker = New frm_colorPicker(strColourPickerMode)
                    'frmPicker.Show()


                Case "grpTbls_Custom"
                    rng = objGlobals.glb_get_wrdSelRng()
                    If rng.Tables.Count <> 0 Then
                        MsgBox("Your cursor must be located at least one clear paragraph away from any existing tables, otherwise they'll merge in unexpected ways." + vbCrLf + vbCrLf + "Please relocate your insertion point and try again")
                    Else
                        frm = New frm_TableBuilder()
                        frm.Show()
                    End If
                    'dlg_InsertTable.Show()
                    'tbl = objGlobals.glb_get_wrdSelTbl()
                    'objPlHTable.Plh_Table_Convert_ToAATable(tbl)

                Case "grpTbls_TableColumnHeadingsStyle", "grpTbls_TableColumnHeadingsStyle_small"
                    Call objStylesMgr.applyStyle_TableColumnHeadings()
                Case "grpTbls_TableUnitsRowStyle", "grpTbls_TableUnitsRowStyle_small"
                    Call objStylesMgr.applyStyle_TableUnitsRow()
                Case "grpTbls_TableTextStyle"
                    Call objStylesMgr.applyStyle_TableText()
                Case "grpTbls_TableTextStyle_small"
                    Call objStylesMgr.applyStyle_TableText_small()
                '
                Case "grpTbls_TableListBullet"
                    Call objStylesMgr.applyStyle_TableListBullet()
                Case "grpTbls_TableListBullet_small"
                    Call objStylesMgr.applyStyle_TableListBullet_small()
                Case "grpTbls_TableListBullet2"
                    Call objStylesMgr.applyStyle_TableListBullet2()
                Case "grpTbls_TableListBullet2_small"
                    Call objStylesMgr.applyStyle_TableListBullet2_small()
                Case "grpTbls_TableListBullet3"
                    Call objStylesMgr.applyStyle_TableListBullet3()
                Case "grpTbls_TableListBullet3_small"
                    Call objStylesMgr.applyStyle_TableListBullet3_small()
                '
                Case "grpTbls_ListNumber"
                    Call objStylesMgr.applyStyle_TableListNumber()
                Case "grpTbls_ListNumber_small"
                    Call objStylesMgr.applyStyle_TableListNumber_small()
                Case "grpTbls_ListNumber2"
                    Call objStylesMgr.applyStyle_TableListNumber2()
                Case "grpTbls_ListNumber2_small"
                    Call objStylesMgr.applyStyle_TableListNumber2_small()
                Case "grpTbls_ListNumber3"
                    Call objStylesMgr.applyStyle_TableListNumber3()
                Case "grpTbls_ListNumber3_small"
                    Call objStylesMgr.applyStyle_TableListNumber3_small()

                Case "grpTbls_Quote"
                    Call objStylesMgr.applyStyle_TableQuote()
                Case "grpTbls_QuoteListBullet"
                    Call objStylesMgr.applyStyle_TableQuoteBullet()
                Case "grpTbls_QuoteSource"
                    Call objStylesMgr.applyStyle_TableQuoteSource()
                'Call objStylesMgr.applyStyle_TableListNumber3_small()

                Case "grpTbls_Quote_small"
                    Call objStylesMgr.applyStyle_TableQuote_small()
                Case "grpTbls_QuoteListBullet_small"
                    Call objStylesMgr.applyStyle_TableQuoteBullet_small()
                Case "grpTbls_QuoteSource_small"
                    Call objStylesMgr.applyStyle_TableQuoteSource_small()

                Case "grpTbls_TableSideHeading1"
                    Call objStylesMgr.applyStyle_TableSideHeading1()
                Case "grpTbls_TableSideHeading1_small"
                    Call objStylesMgr.applyStyle_TableSideHeading1_small()
                Case "grpTbls_TableSideHeading2"
                    Call objStylesMgr.applyStyle_TableSideHeading2()
                Case "grpTbls_TableSideHeading2_small"
                    Call objStylesMgr.applyStyle_TableSideHeading2_small()
                '
                Case "grpTbls_StyleSet_TableQuote"
                    rng = objGlobals.glb_get_wrdApp.Selection.Range
                    objStylesMgr.insert_ExampleTableText_column01(rng, "normal")
                Case "grpTbls_StyleSet_TableListBullets"
                    rng = objGlobals.glb_get_wrdApp.Selection.Range
                    objStylesMgr.insert_ExampleTableText_column02(rng, "normal")
                Case "grpTbls_StyleSet_TableListNumbers"
                    rng = objGlobals.glb_get_wrdApp.Selection.Range
                    objStylesMgr.insert_ExampleTableText_column03(rng, "normal")
                '
                Case "grpTbls_StyleSet_TableQuote_small"
                    rng = objGlobals.glb_get_wrdApp.Selection.Range
                    objStylesMgr.insert_ExampleTableText_column01(rng, "small")
                Case "grpTbls_StyleSet_TableListBullets_small"
                    rng = objGlobals.glb_get_wrdApp.Selection.Range
                    objStylesMgr.insert_ExampleTableText_column02(rng, "small")
                Case "grpTbls_StyleSet_TableListNumbers_small"
                    rng = objGlobals.glb_get_wrdApp.Selection.Range
                    objStylesMgr.insert_ExampleTableText_column03(rng, "small")
                '
                Case "grpTbls_ColourCells"
                    '
                    If objGlobals.glb_get_wrdSel.Range.Tables.Count = 0 Then
                        MsgBox("Make certain that you have selected the cells that you wish to fill")
                        Exit Sub
                    End If
                    '
                    objTables.tbl_colour_set_colourOfCells(objGlobals.glb_get_wrdSel.Cells, objTables._glb_colour_UnitsGrey)

                Case "grpTbls_ColourHeadingsRow"
                    '
                    If objGlobals.glb_get_wrdSelRngAll.Tables.Count = 0 Then
                        Call objMsgMgr.colourRowsErrorMessage()
                        Exit Sub
                    End If
                    '
                    rng = objGlobals.glb_get_wrdSelRngAll
                    If rng.Rows.Count = 0 Then
                        Call objMsgMgr.colourRowsErrorMessage()
                        Exit Sub
                    End If
                    '
                    For Each dr In rng.Rows
                        Call objTables.tbl_colour_set_colourOfRow(dr, objTables._glb_colour_purple_Dark)
                    Next dr
            '
                Case "grpTbls_ColourUnitsRow"
                    If objGlobals.glb_get_wrdSelRngAll.Tables.Count = 0 Then
                        Call objMsgMgr.colourRowsErrorMessage()
                        Exit Sub
                    End If
                    '
                    rng = objGlobals.glb_get_wrdApp.Selection.Range
                    If rng.Rows.Count = 0 Then
                        Call objMsgMgr.colourRowsErrorMessage()
                        Exit Sub
                    End If
                    For Each dr In rng.Rows
                        Call objTables.tbl_colour_set_colourOfRow(dr, objTables._glb_colour_UnitsGrey)
                    Next dr
                Case "grpTbls_AllBorders"
                    rng = objGlobals.glb_get_wrdSelRng
                    '
                    objTables.tbl_borders_colourAndVisibility(rng, True, objTables._glb_colour_TableBorders)
            '
                Case "grpTbls_AllBordersRemove"
                    rng = objGlobals.glb_get_wrdSelRng
                    '
                    For Each tbl In objGlobals.glb_get_wrdSelRngAll.Tables
                        Call objTables.tbl_doBorders_MaintainPadding(tbl, False, objTables._glb_colour_TableBorders)
                    Next tbl
                    '
                Case "grpTbls_ApplyBottomBorder"
                    MsgBox("Still necessary, given the current table facilities..?")

                Case "grpTbls_convertTabletoES"
                    objPlHTable.PlhTbl_Captions_ConvertCaptionsTo_ES(objGlobals.glb_get_wrdSelRngAll)
                    objFlds.updateSequenceNumbers_Tables_ES()

                Case "grpTbls_convertTabletoStd"
                    objPlHTable.PlhTbl_Captions_ConvertCaptionsTo_Report(objGlobals.glb_get_wrdSelRngAll)
                    objFlds.updateSequenceNumbers_Tables()

                Case "grpTbls_convertTabletoApp"
                    objPlHTable.PlhTbl_Captions_ConvertCaptionsTo_Appendix(objGlobals.glb_get_wrdSelRngAll)
                    objFlds.updateSequenceNumbers_Tables_AP()

        '*******
                Case "grpTbls_convertTabletoLT"
                    objPlHTable.PlhTbl_Captions_ConvertCaptionsTo_Letter(objGlobals.glb_get_wrdSelRngAll)
                    objFlds.updateSequenceNumbers_Tables_LT()

                Case "grpTbls_convertTabletoX"
                Case Else
            End Select

        Catch ex As Exception

        End Try
        '
        objGlobals.glb_get_wrdApp.ScreenUpdating = True
        '
    End Sub
    '

    Private Sub PIF_Plh_grpCaptionsMgmnt_Click(sender As Object, e As RibbonControlEventArgs) Handles mnu_grpViewTools_Refresh_btn_setRefFldNotBold.Click, grp_Finalise_CrossRefFlds_setToArialNarrow.Click, grp_Finalise_CrossRefFlds_setRefFldNotBold.Click
        Dim dlgRslt As Integer
        Dim rng As Word.Range
        Dim objGlobals As New cGlobals()
        Dim objToolsMgr As cTools
        Dim objCaptionsMgr As cCaptionManager
        Dim oldRange As Word.Range
        Dim objTOCMgr As New cTOCMgr()
        '
        'oldApplicationScreenUpdating = Globals.ThisDocument.Application.ScreenUpdating
        oldRange = objGlobals.glb_get_wrdSel().Range
        '
        objToolsMgr = New cTools()
        objCaptionsMgr = New cCaptionManager()
        '
        'Convert Captions.. Early version had a '-' separated caption this needed to be
        'chnaged to a tab separated caption in order to facilitate tabbing in the TOC.
        'This function is only necessary to modify legacy documents. It's lifetime is short
        '
        Select Case e.Control.Id
            Case "mnu_grpCaptionsMgmnt_btn_insertTab"
                '
                dlgRslt = MsgBox("This function will convert all Table/Figure/Box Captions from '-' separated" & vbCr _
                & " to tab separated. If there are no '-' separated Captions, nothing happens.. " & vbCr & vbCr _
                & "Be patient, this could take up to a minute on a large document running on a slow machine." & vbCr & vbCr _
                & "A dialogue box will appear onscreen when the Caption conversion process is complete...." _
                & "This action cannot be undone" & vbCr & vbCr _
                & "Do you wish to continue.?", vbYesNo + vbDefaultButton2, "Caption Conversion Warning")
                If dlgRslt = Constants.vbYes Then
                    rng = objGlobals.glb_get_wrdSel()
                    'Set objTempFixes = New cTempFixes
                    'Call objTempFixes.convertCaptionToTab
                    Call objCaptionsMgr.convertCaptionToTab()
                    rng.Select()
                    'Now update TOC and Table Of Figures
                    Call objTOCMgr.toc_update_TOCs(objGlobals.glb_get_wrdActiveDoc)
                    Call objTOCMgr.toc_upDate_TOFs()
                    MsgBox("Caption conversions have been completed")
                End If
            Case "grp_Finalise_CrossRefFlds_setToArialNarrow"
                MsgBox("This method will scan the entire document and set the font of all 'cross reference' fields to the 'Arial Narrow' standard")
                objCaptionsMgr.cpt_setCrossRef_FieldsToBodyTextFont()

            Case "mnu_grpCaptionsMgmnt_btn_setRefFldNotBold", "grp_Finalise_CrossRefFlds_setRefFldNotBold", "mnu_grpViewTools_Refresh_btn_setRefFldNotBold"
                'Sets the cross reference fields to not bold
                MsgBox("This method will scan the entire document and set all 'cross reference' fields to 'NOT bold'")
                objCaptionsMgr.setFieldsBoldStatus(False)
            Case "mnu_grpCaptionsMgmnt_btn_setRefFldBold", "mnu_grpViewTools_Refresh_btn_setRefFldBold"
                'Resets to Equation/Table/Figure and Table and Figure numbering
                'is set to Heading1.Sequence
                MsgBox("This method will scan the entire document and set all 'cross reference' fields to 'bold'")
                objCaptionsMgr.setFieldsBoldStatus(True)
            Case "mnu_grpCaptionsMgmnt_btn_deleteCaptions"
                'Resets to Equation/Table/Figure and Table and Figure numbering
                'is set to Heading1.Sequence
                Call objCaptionsMgr.deleteAllNotBuiltInCaptions()
            Case "mnu_grpCaptionsMgmnt_btn_resetCaptions"
                Call objCaptionsMgr.deleteAllNotBuiltInCaptions()
                Call objCaptionsMgr.installCustomCaptions()
            Case Else
        End Select
        '
        oldRange.Select()
        '
        'Re-establish the current Screen Updating selection
        'oldApplicationScreenUpdating = Globals.ThisDocument.Application.ScreenUpdating
        '
    End Sub

    Private Sub PIF_PgS_grpReport_Columns_Click(sender As Object, e As RibbonControlEventArgs) Handles grpReport_Columns_04.Click, grpReport_Columns_03.Click, grpReport_Columns_02_RightWide.Click, grpReport_Columns_02_LeftWide.Click, grpReport_Columns_02.Click, grpReport_Columns_01.Click
        Dim objRptMgr As New cReport()
        Dim objSectMgr As New cSectionMgr()
        Dim objColsMgr As New cColsHandler()
        Dim objGlobals As New cGlobals()
        Dim sect As Word.Section
        '
        sect = objGlobals.glb_get_wrdSect()
        objGlobals.glb_screen_update(False)

        Select Case e.Control.Id
            Case "grpReport_Columns_04"
                If objColsMgr.cols_isOK_ToChangeColumns(sect) Then
                    objColsMgr.cols_setup_columnStructure(sect, "4_columns")
                End If

            Case "grpReport_Columns_03"
                If objColsMgr.cols_isOK_ToChangeColumns(sect) Then
                    objColsMgr.cols_setup_columnStructure(sect, "3_columns")
                End If

            Case "grpReport_Columns_02"
                If objColsMgr.cols_isOK_ToChangeColumns(sect) Then
                    objColsMgr.cols_setup_columnStructure(sect, "2_columns")
                End If

            Case "grpReport_Columns_02_LeftWide"
                If objColsMgr.cols_isOK_ToChangeColumns(sect) Then
                    objColsMgr.cols_setup_columnStructure(sect, "2_columns_leftWide")
                End If

            Case "grpReport_Columns_02_RightWide"
                If objColsMgr.cols_isOK_ToChangeColumns(sect) Then
                    objColsMgr.cols_setup_columnStructure(sect, "2_columns_rightWide")
                End If

            Case "grpReport_Columns_01"
                If objColsMgr.cols_isOK_ToChangeColumns(sect) Then
                    objColsMgr.cols_setup_columnStructure(sect, "1_columns")
                End If

        End Select
        '
        objGlobals.glb_screen_update(True)
        '
    End Sub

    Private Sub PIF_PgS_grpAppendix_Click(sender As Object, e As RibbonControlEventArgs) Handles grpAppendix_newAttPart.Click, grpAppendix_newAppChapter_inFront_bblk.Click, grpAppendix_newAppChapter_behind_bblk.Click, grpAppendix_newAppChapter_inFront.Click, grpAppendix_newAppChapter_behind.Click
        Dim objSectMgr As New cSectionMgr()
        Dim objChpt As New cChptApp()
        Dim objDivMgr As cChptDivider
        Dim objRptMgr As New cReport()
        Dim objWrkAround As New cWorkArounds()
        Dim objStylesMgr As New cStylesManager()
        Dim objisOKMgr As New cIsOKToDo()
        Dim objBBlkMgr As New cBBlocksHandler()
        Dim objViewMgr As New cViewManager()
        Dim sect As Word.Section
        Dim rng As Word.Range
        Dim placeBehind As Boolean
        Dim strHeading, strRptMode, strErrorMsg As String
        '
        strErrorMsg = ""
        sect = objSectMgr.objGlobals.glb_get_wrdApp.Selection.Sections.Item(1)
        objSectMgr.objGlobals.glb_screen_update(False)
        '
        placeBehind = True
        strRptMode = objRptMgr.Rpt_Mode_Get()
        '
        Select Case e.Control.Id
            Case "grpAppendix_newAppPart"
                objDivMgr = New cChptDivider()
                'strHeading = "Appendices" + vbCr + "Sub heading"
                'objDivMgr.div_insert_newAP(Not placeBehind, sect, strRptMode)
                Select Case strRptMode
                    Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                        objDivMgr.div_insert_newAP(Not placeBehind, sect, strRptMode)
                    Case objRptMgr.rpt_isBrief
                        MsgBox("This function is not available in an ACIL Allen Brief.")
                End Select
            Case "grpAppendix_newAttPart"
                objDivMgr = New cChptDivider()
                strHeading = "Attachment"
                Select Case strRptMode
                    Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                        objDivMgr.div_insert_newAP(Not placeBehind, sect, strRptMode, strHeading)
                    Case objRptMgr.rpt_isBrief
                        MsgBox("This function is not available in an ACIL Allen Brief.")
                End Select

            Case "grpAppendix_newAppChapter_inFront_bblk"
                strErrorMsg = objisOKMgr.isOKto_Insert_ChptAPinFront(sect, strRptMode)
                '
                objViewMgr.vw_change_ColumnsAndRows(sect)
                '
                If strErrorMsg = "" Then
                    rng = objBBlkMgr.bblk_insert_sectionInFront("aa_Chpt_Ap")
                    rng.Select()
                    '
                Else
                    MsgBox(strErrorMsg)
                End If
                '
            Case "grpAppendix_newAppChapter_behind_bblk"
                strErrorMsg = objisOKMgr.isOKto_Insert_ChptAPBehind(sect, strRptMode)
                objViewMgr.vw_change_ColumnsAndRows(sect)
                '
                Try
                    If strErrorMsg = "" Then
                        Dim myDoc As Word.Document
                        myDoc = sect.Range.Document
                        sect = objSectMgr.objGlobals.glb_get_wrdSect
                        If Not (sect.Index = myDoc.Sections.Last.Index) Then
                            sect = myDoc.Sections.Item(sect.Index + 1)
                            rng = sect.Range
                            rng.Collapse(WdCollapseDirection.wdCollapseStart)
                            rng.Select()
                            rng = objBBlkMgr.bblk_insert_sectionInFront("aa_Chpt_Ap")
                            rng.Select()
                        Else
                            MsgBox("Inserting behind the last section is not permitted")
                        End If
                        '
                    Else
                        MsgBox(strErrorMsg)
                    End If

                Catch ex As Exception

                End Try

            Case "grpAppendix_newAppChapter_inFront"
                strErrorMsg = objisOKMgr.isOKto_Insert_ChptAPinFront(sect, strRptMode)
                '
                If strErrorMsg = "" Then
                    rng = objChpt.app_insert_App(Not placeBehind, sect, strRptMode)
                    rng = objChpt.chptBase_getRange_Heading1(rng)
                    rng.Select()
                    '
                Else
                    MsgBox(strErrorMsg)
                End If

                'tbl = objChpt.app_insert_App(Not placeBehind, sect, strRptMode)
                'objChpt.chptBase_select_Chapter(tbl)
            Case "grpAppendix_newAppChapter_behind"
                strErrorMsg = objisOKMgr.isOKto_Insert_ChptAPBehind(sect, strRptMode)
                '
                If strErrorMsg = "" Then
                    rng = objChpt.app_insert_App(placeBehind, sect, strRptMode)
                    '
                    rng = objChpt.chptBase_getRange_Heading1(rng)
                    rng.Select()
                    'objChpt.chptBase_select_Chapter(rng)
                Else
                    MsgBox(strErrorMsg)
                End If


            Case "grpAppendix_Landscape"
                sect = objChpt.app_insert_landscapeSection()
                rng = sect.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Select()

        End Select
        '
        objWrkAround.wrk_fix_forCursorRace()
        '
        objSectMgr.objGlobals.glb_screen_update(True)

    End Sub

    Private Sub PIF_PgS_grpToc_Click(sender As Object, e As RibbonControlEventArgs) Handles grpToc_TOC_insertSection.Click, grpToc_TOC_insertLevels_1_to_3.Click, grpToc_TOC_insertLevels_1_to_2.Click, grpToc_TOC_insertLevels_1_to_1.Click
        Dim objSectMgr As New cSectionMgr()
        Dim objTOCMgr As New cTOCMgr()
        Dim objRptMgr As New cReport()
        Dim objGlobals As New cGlobals()
        Dim objMsgMgr As New cMessageManager()
        Dim para As Word.Paragraph
        Dim sect As Word.Section
        Dim rng As Word.Range
        Dim placeBehind As Boolean
        Dim strRptMode As String
        '
        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = False
        sect = objSectMgr.objGlobals.glb_get_wrdSelRng.Sections.Item(1)
        '
        placeBehind = False
        strRptMode = objRptMgr.Rpt_Mode_Get()
        '
        Try
            '
            Select Case e.Control.Id
                Case "grpToc_TOC_insertSection"
                    Select Case strRptMode
                        Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                            If Not objTOCMgr.toc_has_TOCSection(sect.Range.Document) Then
                                '
                                sect = objTOCMgr.toc_insert_TOCSection(True, placeBehind, sect, strRptMode)
                                rng = sect.Range()
                                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                                rng.MoveStart(WdUnits.wdParagraph, -3)
                                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                                rng.Select()
                                '
                            Else
                                MsgBox("This document already has a Table of Contents Page")
                            End If
                            '
                        Case objRptMgr.rpt_isBrief
                            ' If Not Globals.ThisDocument.Application.Selection.Information(WdInformation.wdWithInTable) Then
                            If Not objSectMgr.sct_Sel_IsIn_TableOnly() Then
                                'If Not objSectMgr.sct_Sel_IsIn_Or_JustUnderTable() Then
                                rng = objGlobals.glb_get_wrdSelRng
                                para = rng.Paragraphs.Add()
                                para.Style = rng.Document.Styles.Item("TOC Heading Brief")
                                rng = para.Range
                                'rng.Font.Color = RGB(108, 61, 153)
                                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                                rng.Text = "Contents"
                                rng.Move(WdUnits.wdParagraph, 1)
                                rng.Collapse(WdCollapseDirection.wdCollapseStart)

                                'objTOCMgr.toc_replace_TOCField(objGlobals.glb_get_wrdSelRng, "aac_TOC_Levels02")
                                objTOCMgr.toc_replace_TOCField(rng, "aac_TOC_Levels02")
                            Else
                                objMsgMgr.msg_insertionPoint_isInsideATable()
                            End If
                    End Select

                Case "grpToc_TOC_insertLevels_1_to_1"
                    If objTOCMgr.toc_replace_TOCField_Levels("aac_TOC_Levels01") Then

                    Else
                        MsgBox("For this to work your cursor needs to be in the current TOC.")
                    End If
                Case "grpToc_TOC_insertLevels_1_to_2"
                    If objTOCMgr.toc_replace_TOCField_Levels("aac_TOC_Levels02") Then

                    Else
                        MsgBox("For this to work your cursor needs to be in the current TOC.")
                    End If
                Case "grpToc_TOC_insertLevels_1_to_3"
                    If objTOCMgr.toc_replace_TOCField_Levels("aac_TOC_Levels03") Then

                    Else
                        MsgBox("For this to work your cursor needs to be in the current TOC.")
                    End If
                '

                Case "grpToc_TOC_wide"
                Case "grpToc_TOC_narrow"

            End Select
            '
        Catch ex As Exception

        End Try
        '
        objGlobals.glb_screen_update(True)
        '
    End Sub

    Private Sub PIF_PgS_grpContactsPages_Click(sender As Object, e As RibbonControlEventArgs) Handles grpContactsPages_ProposalTo.Click, grpContactsPages_FrontPage_AckOfCountry.Click, grpContactsPages_FrontPage.Click, grpContactsPages_BackPage.Click, grpContactsPages_Disclaimer.Click, grpContactsPages_CopyrightStatement.Click, grpContactsPages_ReportTo.Click
        Dim objSectMgr As New cSectionMgr()
        Dim objContactsMgr As New cContactsMgr()
        Dim objRptMgr As New cReport()
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objBnrMgr As New cChptBanner()
        Dim objMsgMgr As New cMessageManager()
        Dim sect As Word.Section
        Dim rng As Word.Range
        Dim objGlobals As New cGlobals()
        Dim dlgResult As Integer
        Dim placeBehind As Boolean
        Dim strRptMode As String
        Dim doBottomOfPageImage As Boolean
        Dim myDoc As Word.Document
        Dim strOrientation As String
        '
        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = False
        placeBehind = True
        doBottomOfPageImage = True      'Report default position
        myDoc = objGlobals.glb_get_wrdActiveDoc
        '
        strOrientation = "prt"
        '
        Try
            strRptMode = objRptMgr.Rpt_Mode_Get()
            sect = objSectMgr.objGlobals.glb_get_wrdSelRng.Sections.Item(1)
            '
            If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then strOrientation = "lnd"
            '
            'If strRptMode = objRptMgr.modeLongLandscape Then strOrientation = "lnd"
            '
            Select Case e.Control.Id
                Case "grpContactsPages_FrontPage_AckOfCountry"
                    doBottomOfPageImage = True
                    '
                    Select Case strRptMode
                        Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                            objContactsMgr.contacts_insert_FrontPage(Not placeBehind, sect, strRptMode, doBottomOfPageImage)
                            '
                        Case objRptMgr.rpt_isBrief
                            MsgBox(objMsgMgr.msgMgr_msg_notAvailableInBrief())
                            '
                    End Select
                    '
                Case "grpContactsPages_FrontPage"
                    doBottomOfPageImage = False
                    '
                    Select Case strRptMode
                        Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                            objContactsMgr.contacts_insert_FrontPage(Not placeBehind, sect, strRptMode, doBottomOfPageImage)
                            '
                        Case objRptMgr.rpt_isBrief
                            MsgBox(objMsgMgr.msgMgr_msg_notAvailableInBrief())
                            '
                    End Select
                    '
                Case "grpContactsPages_BackPage"
                    'objContactsMgr.contacts_insert_BackPage()
                    If objContactsMgr.has_ContactsPage_Back() Then
                        MsgBox("This document already has a back Contacts Page")
                    Else
                        rng = objGlobals.glb_get_wrdSel.Range
                        dlgResult = MsgBox("You are about To insert a Back Page" & vbCr & vbCr _
                    & "Do you wish To Continue With the back page insertion.?", vbYesNo + vbDefaultButton2, "Back Page Warning")
                        If dlgResult = vbYes Then
                            '
                            sect = objContactsMgr.contacts_insert_BackPage()
                            objHfMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_cont_Back))
                            objHfMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_cont_Back), "firstPage")
                            '
                            GoTo loop4



                            'Call objContactsPageMgr.insert_ContactsPage_Back()
                            rng = myDoc.Sections.Last.Range
                            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                            sect = myDoc.Sections.Add(rng, WdSectionStart.wdSectionNewPage)
                            sect = myDoc.Sections.Last
                            '
                            If strOrientation = "prt" Then sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait
                            If strOrientation = "lnd" Then sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape
                            '
                            sect.PageSetup.DifferentFirstPageHeaderFooter = True
                            objHfMgr.hf_hfs_linkUnlinkAll(sect, False)
                            rng = sect.Range
                            rng.Collapse(WdCollapseDirection.wdCollapseStart)
                            rng.Paragraphs.Add(rng)
                            rng.Paragraphs.Add(rng)
                            rng.Paragraphs.Add(rng)
                            rng.Paragraphs.Add(rng)
                            rng.Collapse(WdCollapseDirection.wdCollapseStart)
                            '
                            'sect = objGlobals.glb_get_wrdSect()
                            '
                            objContactsMgr.contacts_convert_toBackContacts(myDoc.Sections.Last)
                            'objContactsMgr.contacts_insert_BackPage()
                            objHfMgr.hf_tags_setTagStyle(myDoc.Sections.Last, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_cont_Back))
                            objHfMgr.hf_tags_setTagStyle(myDoc.Sections.Last, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_cont_Back), "firstPage")
loop4:
                            'rng.Select()
                        End If
                    End If

                Case "grpContactsPages_ReportTo"
                    objContactsMgr.contacts_insert_Citation("front", "report_to")
                Case "grpContactsPages_ProposalTo"
                    objContactsMgr.contacts_insert_Citation("front", "proposal_to")

                Case "grpContactsPages_CopyrightStatement"
                    objContactsMgr.contacts_insert_Citation("front", "copyrightStatement")
                    'objContactsMgr.conts_insert_CitationContacts("front", "proposal")
                Case "grpContactsPages_Disclaimer"
                    objContactsMgr.contacts_insert_Citation("front", "disclaimer")
                    'objContactsMgr.conts_insert_CitationContacts("front", "proposal")

            End Select
            '
        Catch ex As Exception

        End Try
        '
        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = True
        '
    End Sub

    Private Sub PIF_PgS_grpCpImages_Click(sender As Object, e As RibbonControlEventArgs) Handles grpCpImages_ImageFromFile.Click, grpCpImages_ImageFromClip.Click, grpCpImages_BackPanelFill_RawImageFromFile.Click, grpCpImages_Delete_SmallPict.Click
        Dim objImgGetEdit As New cImageGetAndEdit()

        Dim objSectMgr As New cSectionMgr()
        Dim objCpMgr As New cCoverPageMgr()
        Dim objImgMgr As New cImageMgr()
        Dim objPanelMgr As New cBackPanelMgr()
        Dim objMsgMgr As New cMessageManager()
        Dim objCropMgr As New cCropMgr()
        Dim objBackPanel As New cShapeMgr()
        Dim objWrkArounds As New cWorkArounds()
        Dim strFilePath As String

        Dim objGlobals As New cGlobals()
        Dim myDoc As Word.Document
        Dim sect As Word.Section
        Dim objViewMgr As New cViewManager()
        '
        'objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = False
        strFilePath = ""
        '
        Try
            myDoc = objSectMgr.objGlobals.glb_get_wrdActiveDoc()
            sect = objSectMgr.objGlobals.glb_get_wrdSect()
            '
            Select Case e.Control.Id
                Case "grpCpImages_ImageFromFile"
                    objImgGetEdit.imgGet_fill_backPanelFromFile_cp_pict_large()
                    'objViewMgr.vw_change_toPageFitBestFit(objSectMgr.objGlobals.glb_get_wrdSect())
                    '
                Case "grpCpImages_ImageFromClip"
                    '
                    objImgGetEdit.imgGet_fill_backPanelFromClipboard_cp_pict_large()
                    '
                Case "grpCpImages_BackPanelFill_RawImageFromFile"
                    '
                    objImgGetEdit.imgGet_fill_withRawImageFromFile_cp_pict_large()
                    '
                Case "grpCpImages_Reset_backcolour"
                    'objImgMgr = New cImageMgr()
                    'objPanelMgr = New cBackPanelMgr()
                    '
                    'Check for back panels
                    'sect = Globals.ThisDocument.Application.Selection.Sections.Item(1)
                    'lstOfBackPanels = objPanelMgr.pnl_getBackPanel_PlaceHolders(sect)                     'To get rid of any existing back panels
                    '
                    'If lstOfBackPanels.Count > 0 Then
                        'objShpMgr = lstOfBackPanels.Item(0)
                        'objPanelMgr.pnl_reset_BackPanelColour(objShpMgr)
                        'Else
                        'MsgBox("Make certain that your cursor is in a" + vbCrLf + "cover page, contacts page or a divider")
                    'End If           
                    '
                Case "grpCpImages_Delete_SmallPict"
                    sect = objGlobals.glb_get_wrdApp.Selection.Sections.Item(1)
                    'Check for pict panels
                    If Not objCpMgr.ChptBase_delete_SmallPicturePlaceHolders(sect) Then
                        MsgBox("Function complete.. If nothing happened, it's probably" + vbCrLf + "because there is no 'small Cover Page picture panel" + vbCrLf + vbCrLf + "Was your cursor in a Cover Page containing a small picture panel?")
                    End If
                    '
                Case Else
                    '
            End Select
            '
        Catch ex As Exception

        End Try
        '
        'PIF_PgS_grpCpImages
        objSectMgr.objGlobals.glb_screen_update(True)
        '
        objWrkArounds.wrk_fix_forCursorRace()
        '
    End Sub

    Private Sub PIF_Plh_grpBoxes_Click(sender As Object, e As RibbonControlEventArgs) Handles grpBoxes_Box.Click, grpBoxes_LTBox.Click, grpBoxes_ESBox.Click, grpBoxes_CaptionAndHeadingES.Click, grpBoxes_CaptionAndHeadingApp.Click, grpBoxes_CaptionAndHeading.Click, grpBoxes_AppendixBox.Click, grpBoxes_SideHeading2.Click, grpBoxes_SideHeading1.Click, grpBoxes_BoxTextBoldItalic.Click, grpBoxes_BoxText.Click, grpBoxes_BoxListNumber3.Click, grpBoxes_BoxListNumber2.Click, grpBoxes_BoxListNumber.Click, grpBoxes_BoxListBullet3.Click, grpBoxes_BoxListBullet2.Click, grpBoxes_BoxListBullet.Click, grpBoxes_BoxQuoteSource.Click, grpBoxes_BoxQuoteListBullet.Click, grpBoxes_BoxQuote.Click, grpBoxes_fillWithExampleStyles.Click, grpBoxes_deleteBoxContent.Click, grpBoxes_ToLT.Click, grpBoxes_ToES.Click, grpBoxes_ToBox1.Click, grpBoxes_ToApp.Click, grpBoxes_RecommendationES.Click, grpBoxes_Recommendation.Click, grpBoxes_KeyFindingES.Click, grpBoxes_KeyFinding.Click
        Dim objWrkAround As New cWorkArounds()
        Dim objSectMgr As New cSectionMgr()
        Dim objRptMgr As New cReport()
        Dim objPlHMgr As New cPlHBase()
        Dim objPlHBox As New cPlhBox()
        Dim objStylesMgr As New cStylesManager()
        Dim objFldsMgr As New cFieldsMgr()
        Dim sect As Word.Section
        Dim objGlobals As New cGlobals()
        Dim placeBehind As Boolean
        Dim strRptMode, strMsg As String
        Dim tbl As Word.Table
        Dim rng As Word.Range
        '
        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = False
        strMsg = ""
        placeBehind = True
        '
        Try
            strRptMode = objRptMgr.Rpt_Mode_Get()
            sect = objSectMgr.objGlobals.glb_get_wrdSelRng.Sections.Item(1)
            '
            Select Case e.Control.Id
                Case "grpBoxes_Box"
                    tbl = objPlHBox.PlhBox_insert_Box("Box")
                Case "grpBoxes_AppendixBox"
                    tbl = objPlHBox.PlhBox_insert_Box("Box_AP")
                Case "grpBoxes_ESBox"
                    tbl = objPlHBox.PlhBox_insert_Box("Box_ES")
                Case "grpBoxes_LTBox"
                    tbl = objPlHBox.PlhBox_insert_Box("Box_LT")
                Case "grpBoxes_KeyFinding"
                    tbl = objPlHBox.PlhBox_insert_Box("Key_Finding")
                    objSectMgr.objGlobals.glb_screen_update()

                    'objWrkAround.wrk_fix_forCursorRace()

                Case "grpBoxes_KeyFindingES"
                    tbl = objPlHBox.PlhBox_insert_Box("Key_Finding_ES")
                    objSectMgr.objGlobals.glb_screen_update()

                    'objWrkAround.wrk_fix_forCursorRace()

                Case "grpBoxes_Recommendation"
                    tbl = objPlHBox.PlhBox_insert_Box("Recommendation")
                Case "grpBoxes_RecommendationES"
                    tbl = objPlHBox.PlhBox_insert_Box("Recommendation_ES")
                Case "grpBoxes_RecommendationLT"
                    tbl = objPlHBox.PlhBox_insert_Box("Recommendation_LT")
                Case "grpBoxes_CaptionAndHeading"
                    rng = objGlobals.glb_get_wrdSelRng()
                    rng = objPlHMgr.Plh_Captions_InsertCaptions("Box", rng, True)
                    rng.Select()
                Case "grpBoxes_CaptionAndHeadingES"
                    rng = objGlobals.glb_get_wrdSelRng()
                    rng = objPlHMgr.Plh_Captions_InsertCaptions("Box_ES", rng, True)
                    rng.Select()
                Case "grpBoxes_CaptionAndHeadingApp"
                    rng = objGlobals.glb_get_wrdSelRng()
                    rng = objPlHMgr.Plh_Captions_InsertCaptions("Box_AP", rng, True)
                    rng.Select()
                    '
                Case "grpBoxes_BoxTextBoldItalic"
                    objStylesMgr.applyStyleToSelection("Box Text (Bold Italic)")
                Case "grpBoxes_BoxText"
                    objStylesMgr.applyStyleToSelection("Box Text")
                Case "grpBoxes_SideHeading1"
                    objStylesMgr.applyStyleToSelection("Box Side Heading 1")
                Case "grpBoxes_SideHeading2"
                    objStylesMgr.applyStyleToSelection("Box Side Heading 2")
                Case "grpBoxes_BoxListBullet"
                    objStylesMgr.applyStyleToSelection("Box List Bullet")
                Case "grpBoxes_BoxListBullet2"
                    objStylesMgr.applyStyleToSelection("Box List Bullet 2")
                Case "grpBoxes_BoxListBullet3"
                    objStylesMgr.applyStyleToSelection("Box List Bullet 3")
                Case "grpBoxes_BoxListNumber"
                    objStylesMgr.applyStyleToSelection("Box List Number")
                Case "grpBoxes_BoxListNumber2"
                    objStylesMgr.applyStyleToSelection("Box List Number 2")
                Case "grpBoxes_BoxListNumber3"
                    objStylesMgr.applyStyleToSelection("Box List Number 3")
                Case "grpBoxes_BoxQuote"
                    objStylesMgr.applyStyleToSelection("Box Quote")
                Case "grpBoxes_BoxQuoteListBullet"
                    objStylesMgr.applyStyleToSelection("Box Quote List Bullet")
                Case "grpBoxes_BoxQuoteSource"
                    objStylesMgr.applyStyleToSelection("Box Quote Source")
                Case "grpBoxes_BoxSourceForOverTyping"
                    'Example text for overtyping, comes in from Building Blocks... Replaced in 2021 by software build

                Case "grpBoxes_ToES"
                    objPlHBox.PlhBox_Captions_ConvertBoxCaptionsTo_ES(objGlobals.glb_get_wrdSel.Range)
                    objFldsMgr.updateSequenceNumbers_Boxes_ES()
                    '
                Case "grpBoxes_ToBox1"
                    objPlHBox.PlhBox_Captions_ConvertBoxCaptionsTo_Report(objGlobals.glb_get_wrdSel.Range)
                    objFldsMgr.updateSequenceNumbers_Boxes()
                    '
                Case "grpBoxes_ToApp"
                    objPlHBox.PlhBox_Captions_ConvertBoxCaptionsTo_Appendix(objGlobals.glb_get_wrdSelRngAll)
                    objFldsMgr.updateSequenceNumbers_Boxes_Ap()

                Case "grpBoxes_ToLT"
                    objPlHBox.PlhBox_Captions_ConvertBoxCaptionsTo_Letter(objGlobals.glb_get_wrdApp.Selection.Range)
                    objFldsMgr.updateSequenceNumbers_Boxes_LT()

                Case "grpBoxes_deleteBoxContent"
                    objPlHBox.PlhBox_Exmaples_InsertBoxText(True)
                'Call objPlhMgr.deleteAndReplaceContent(True)
                Case "grpBoxes_fillWithExampleStyles"
                    'objStylesMgr.insertStyleSetReport_Box()
                    objPlHBox.PlhBox_Exmaples_InsertBoxText(False)
                    'Call objPlhMgr.deleteAndReplaceContent(False)


            End Select
            '
        Catch ex As Exception

        End Try
        '
        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = True
        '
    End Sub

    Private Sub PIF_Plh_grpFigures_Click(sender As Object, e As RibbonControlEventArgs) Handles grpFigures_Figure.Click, grpFigures_StyleForSubHeadings.Click, grpFigures_LT.Click, grpFigures_ES.Click, grpFigures_CaptionAndHeadingES.Click, grpFigures_CaptionAndHeadingApp.Click, grpFigures_CaptionAndHeading.Click, grpFigures_Appendix.Click, grpFigures_convertToStd.Click, grpFigures_convertToLT.Click, grpFigures_convertToES.Click, grpFigures_convertToApp.Click
        Dim rng As Range
        Dim objStylesMgr As New cStylesManager()
        Dim objFldsMgr As New cFieldsMgr()
        Dim sect As Word.Section
        Dim objFigMgr As New cPlhFigure()
        Dim objSectMgr As New cSectionMgr()
        Dim tbl As Word.Table
        '
        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = False
        '
        sect = objSectMgr.objGlobals.glb_get_wrdSect
        '
        Try
            Select Case e.Control.Id
                Case "grpTblsPlh_ToggleFigWidth"
                    'tbl = objFigMgr.glb_get_wrdSelTbl()
                    'If Not IsNothing(tbl) Then
                    'objFigMgr.objGlobals.glb_tbls_BannerAutoFit(tbl)
                    'Else
                    'MsgBox("For this method to work your cursor needs to be in the Figure that you want to adjust, and that figure must be in the Acil Allen standard format")
                    'End If
                Case "grpFigures_PlaceHolder_4cm"
                Case "grpFigures_Figure"
                    tbl = objFigMgr.PlhFig_insert_Figure("Figure")
                Case "grpFigures_Figure_17p5"
                    objFigMgr.PlhFig_insert_FigureWide("Figure")
                Case "grpFigures_Appendix"
                    objFigMgr.PlhFig_insert_Figure("Figure_AP")
                Case "grpFigures_Appendix_17p5"
                    objFigMgr.PlhFig_insert_FigureWide("Figure")
                Case "grpFigures_ES"
                    objFigMgr.PlhFig_insert_Figure("Figure_ES")
                Case "grpFigures_ES_17p5"
                    objFigMgr.PlhFig_insert_FigureWide("Figure_ES")
                Case "grpFigures_LT"
                    objFigMgr.PlhFig_insert_Figure("Figure_LT")
                Case Else
            End Select
            '
            Select Case e.Control.Id
                Case "grpFigures_CaptionAndHeading"
                    rng = objFigMgr.Plh_Captions_InsertCaptions("Figure", objSectMgr.objGlobals.glb_get_wrdSelRngAll, True)
                    rng.Select()
                Case "grpFigures_CaptionAndHeadingApp"
                    rng = objFigMgr.Plh_Captions_InsertCaptions("Figure_AP", objSectMgr.objGlobals.glb_get_wrdSelRngAll, True)
                    rng.Select()
                Case "grpFigures_CaptionAndHeadingES"
                    rng = objFigMgr.Plh_Captions_InsertCaptions("Figure_ES", objSectMgr.objGlobals.glb_get_wrdSelRngAll, True)
                    rng.Select()
                Case Else
            End Select
            '
            Select Case e.Control.Id
                Case "grpFigures_StyleForSubHeadings"
                    objStylesMgr.applyStyleToSelection("Figure (Sub Headings)")
                    'Patch to force a behaviour without having to re-issue the template. Note that
                    'this was included in the styles definitions of version 12.12.32
                    'For Each para In Globals.ThisDocument.Application.Selection.Range.Paragraphs
                    'para.Format.KeepWithNext = True
                    'para.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                    'Next
                    '
                Case "grpFigures_convertToES"
                    objFigMgr.PlhFig_Captions_ConvertBoxCaptionsTo_ES(objSectMgr.objGlobals.glb_get_wrdSelRngAll)
                    objFldsMgr.updateSequenceNumbers_Figures_ES()
                Case "grpFigures_convertToStd"
                    objFigMgr.PlhFig_Captions_ConvertBoxCaptionsTo_Report(objSectMgr.objGlobals.glb_get_wrdSelRngAll)
                    objFldsMgr.updateSequenceNumbers_Figures()
                Case "grpFigures_convertToApp"
                    objFigMgr.PlhFig_Captions_ConvertBoxCaptionsTo_Appendix(objSectMgr.objGlobals.glb_get_wrdSelRngAll)
                    objFldsMgr.updateSequenceNumbers_Figures_Ap()

        '*******
                Case "grpFigures_convertToLT"
                    objFigMgr.PlhFig_Captions_ConvertBoxCaptionsTo_Letter(objSectMgr.objGlobals.glb_get_wrdSelRngAll)
                    objFldsMgr.updateSequenceNumbers_Figures_LT()
                Case Else
            End Select



        Catch ex As Exception

        End Try
        '
        objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = True

        'Re-establish the current Screen Updating selection
        '
    End Sub

    Private Sub PIF_grpTest_Click(sender As Object, e As RibbonControlEventArgs) Handles grpTest_pgNum_getTagStyleMap.Click, grpTest_btn_cloneDoc.Click, grpTest_btn_getTimeStamp.Click
        'Dim rbnCollection As ThisRibbonCollection
        Dim objGlobals As New cGlobals()
        Dim objSectMgr As cSectionMgr
        Dim objHfMgr As cHeaderFooterMgr
        Dim objGrfxMgr As New cGraphicsMgr()
        Dim objTocMgr As New cTOCMgr()
        Dim objStylesMgr As New cStylesManager()
        Dim objLstStyles As New clstStyles()
        Dim objCloneMgr As New cCloneMgr()
        Dim objkWrkAround As New cWorkArounds()
        Dim objMsgMgr As New cMessageManager()
        Dim lstOfSections As New Collection()
        Dim objFileMgr As cFileHandler
        Dim myDoc, srcDoc, destDoc As Word.Document
        Dim startTime As Date
        Dim stpWatch As System.Diagnostics.Stopwatch
        Dim Interval As TimeSpan
        Dim rng As Word.Range
        Dim sect As Word.Section
        Dim strRslt, strElapsedTime, strFrmMode, strTemplate As String
        'Dim frm As frm_colorPicker
        Dim frmTrans As frm_transparency
        Dim frm_tagStyle_Map As frm_tagStyle_Map
        Dim j As Integer
        '
        'sect = Globals.ThisDocument.Application.Selection.Sections.Item(1)
        sect = objGlobals.glb_get_wrdApp.Selection.Sections.Item(1)
        objGlobals.glb_get_wrdApp.ScreenUpdating = False
        '
        '
        Select Case e.Control.Id
            Case "grpTest_btn_cloneDoc"
                Try
                    If objMsgMgr.msgMgr_dlg_cloneDocumentWarning() Then
                        srcDoc = objGlobals.glb_get_wrdActiveDoc
                        '
                        strTemplate = objGlobals.glb_getTmpl_FullName()
                        destDoc = objGlobals.glb_get_wrdApp.Documents.Add(strTemplate, False, WdNewDocumentType.wdNewBlankDocument, True)
                        objCloneMgr.clone_Doc_byCopy(srcDoc, destDoc)
                        '
                    Else
                        MsgBox("The cloning function has been cancelled by the user")
                    End If
                    '
                Catch ex As Exception

                End Try
            Case "grpTest_btn_getTimeStamp"
                objFileMgr = New cFileHandler()
                strRslt = objFileMgr.file_get_TimeStamp()
                'strRslt = objFileMgr.file_make_dirScratch()
                MsgBox("Time stamp is = " + strRslt)

            Case "grpTest_paras_DetectAndSetTopOfSection"
                MsgBox("OK")
                objSectMgr = New cSectionMgr()
                objSectMgr.sct_set_SelforTableInsert()

            Case "grpTest_pgNum_ribbonTest"
                'Globals.ThisDocument.ribbon.InvalidateControl("cmbBox_test")
                Globals.ThisAddIn.ribbon.InvalidateControl("aac1:cmbBox_test")

                'rbnCollection = Globals.Ribbons
                j = 1
            Case "grpTest_test_listTemplate"
                objStylesMgr.style_apply_ListTemplate(objGlobals.glb_get_wrdSelRngAll)

            Case "grpTest_hfs_deleteHeaders"
                objHfMgr = New cHeaderFooterMgr()
                objHfMgr.hf_headers_delete(sect)
            Case "grpTest_hfs_deleteFooters"
                objHfMgr = New cHeaderFooterMgr()
                objHfMgr.hf_footers_delete(sect)

            Case "grpTest_hfs_insertHeaders"
                objHfMgr = New cHeaderFooterMgr()
                objHfMgr.hf_headers_insert(sect)
            Case "grpTest_hfs_insertFooters"
                objHfMgr = New cHeaderFooterMgr()
                objHfMgr.hf_footers_insert(sect)

            Case "grpTest_hfs_linkToPrevious_No"
                objHfMgr = New cHeaderFooterMgr()
                objHfMgr.hf_hfs_linkUnlinkAll(sect, False)
            Case "grpTest_hfs_linkToprevious_Yes"
                objHfMgr = New cHeaderFooterMgr()
                objHfMgr.hf_hfs_linkUnlinkAll(sect, True)
        End Select
        '
        Select Case e.Control.Id
            Case "grpTest_sect_InsertSection"
                objSectMgr = New cSectionMgr()
                sect = objGlobals.glb_get_wrdSect()
                sect = objSectMgr.sct_insert_Section(False, sect, 6, "newPage", False, "prt,")
                'sect = objSectMgr.sct_insert_SectionInFront(objGlobals.glb_get_wrdApp.Selection.Range,, "oddPage")
                rng = sect.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Select()
            Case "grpTest_sect_InsertSection_Behind"
                objSectMgr = New cSectionMgr()
                'sect = objSectMgr.sct_insert_Section(False, sect, 6, "newPage", False, "lnd,")

                sect = objSectMgr.sct_insert_SectionBehind(objGlobals.glb_get_wrdApp.Selection.Range,, "newPage",, "flow")
                ' rng = sect.Range
                ' rng.Collapse(WdCollapseDirection.wdCollapseStart)
                'rng.Select()
            Case "grpTest_sect_DeleteAllSectionContents"
                objSectMgr = New cSectionMgr()
                objSectMgr.sct_delete_allSectionContents(objGlobals.glb_get_wrdSect)
                '
                objGlobals.glb_screen_updateLeaveAsItWas()


            Case "grpTest_sect_DeleteAll"
                MsgBox("This function will delete all but the last section.. and it will remove all contents from the last section")
                objSectMgr = New cSectionMgr()
                objSectMgr.sct_delete_allSections()

        End Select
        '
        Select Case e.Control.Id
            Case "grpTest_Image_InsertBanner"
                objGrfxMgr.grfx_insert_ImageChapterBanner(objGlobals.glb_get_wrdSelRng, 157)
                'objGrfxMgr.grfx_insertShape_behind(objGlobals.glb_get_wrdSelRng, 397, 157, 28, 3, 61)
            Case "grpTest_Image_Insertxxx"
                objGrfxMgr.grfx_insert_ImageCP(objGlobals.glb_get_wrdSelRng.Sections.Item(1))
        End Select
        '
        Dim objpgNumMgr As New cPageNumberMgr()
        Select Case e.Control.Id
            Case "grpTest_pgNum_2Part"
                objpgNumMgr.pgNum_setBody_numFormat(objGlobals.glb_get_wrdActiveDoc(), "2part")
            Case "grpTest_pgNum_1Part"
                objpgNumMgr.pgNum_setBody_numFormat(objGlobals.glb_get_wrdActiveDoc())
        End Select
        '
        Select Case e.Control.Id
            Case "grpTest_pgHfs_UnlinkAll"
                MsgBox("Got to unlink")
                stpWatch = System.Diagnostics.Stopwatch.StartNew()
                startTime = TimeOfDay()
                '
                objHfMgr = New cHeaderFooterMgr()
                objHfMgr.hf_hfs_UnlinkAllSections(objGlobals.glb_get_wrdActiveDoc())
                '
                stpWatch.Stop()
                Interval = stpWatch.Elapsed()
                '
                strElapsedTime = Int(Interval.TotalSeconds()) & " Seconds"
                '
                objGlobals.glb_screen_update(True)
                MsgBox("All headers and footers are unlinked (" + strElapsedTime + ")")

            Case "grpTest_sect_ReOrientToLnd"
                MsgBox("Got to Re-Orient")
        End Select
        '
        Select Case e.Control.Id
            Case "grpTest_pgNum_getTagStyle"
                objHfMgr = New cHeaderFooterMgr()
                Dim objTagsMgr As New cTagsMgr()
                strRslt = objHfMgr.hf_tags_getTagStyleName(objGlobals.glb_get_wrdSect, "primaryOrFirstPage")
                'strRslt = objTagsMgr.tags_get_tagStyleName(objGlobals.glb_get_wrdSect)
                MsgBox("tag style = " + strRslt)
            Case "grpTest_pgNum_getTagStyleMap"
                '
                'frm_doc_Placeholders = New frm_findTables()
                '
                'frm_doc_Placeholders.StartPosition = FormStartPosition.WindowsDefaultLocation
                'frm_doc_Placeholders.TopMost = True
                'frm_doc_Placeholders.Show()
                '
                '
                'objHfMgr = New cHeaderFooterMgr()
                myDoc = objGlobals.glb_get_wrdActiveDoc()
                '
                frm_tagStyle_Map = New frm_tagStyle_Map(myDoc)
                frm_tagStyle_Map.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
                frm_tagStyle_Map.TopMost = True
                frm_tagStyle_Map.Show()

                '
                'lstOfSections = objHfMgr.hf_getTagStyleMap_All(myDoc)
                'strRslt = ""
                '
                'If lstOfSections.Count > 0 Then
                'strRslt = ""
                'For j = 1 To lstOfSections.Count
                'strRslt = strRslt + lstOfSections.Item(j) + vbCrLf
                'Next
                'End If
                '
                'MsgBox(strRslt, , "Report Tag Style Map")
                '
                '
        End Select
        '
        Select Case e.Control.Id
            Case "grpTest_mnu01_getColorPicker"
                strFrmMode = "testMode"
                objGlobals.glb_show_ColorPicker(strFrmMode)
                'frm = New frm_colorPicker(strFrmMode)
                'frm.Show()
            Case "grpTest_set_Transparency"
                '
                frmTrans = New frm_transparency()
                frmTrans.Show()
                '
                'Globals.Ribbons.ge
                '
                'Me.ribbon.InvalidateControl(control.Id)
                'Me.ribbon.InvalidateControl("grpTest_set_Transparency")

                'If IsNothing(Globals.ThisDocument.frmTrans) Then
                'Globals.ThisDocument.frmTrans = New frm_transparency()
                'Globals.ThisDocument.frmTrans.Show()
                'objkWrkAround.wrk_fix_forCursorRace()
                'Else
                'Globals.ThisDocument.frmTrans.Activate()
                'objkWrkAround.wrk_fix_forCursorRace()
                'End If
        End Select
        '
        objGlobals.glb_get_wrdApp.ScreenUpdating = True

    End Sub

    Private Sub PIF_grpCrypt_Click(sender As Object, e As RibbonControlEventArgs) Handles grpMetaData_Remove_FromDoc.Click
        'Dim frm_cryptTools As frm_cryptTools
        Dim objGlobals As New cGlobals()
        Dim objReportMgr As New cReport()
        Dim objPropsMgr As New cPropertyMgr()
        Dim myDoc As Word.Document
        'Dim myDocInfo As System.IO.FileInfo
        Dim objMsgMgr As New cMessageManager()
        Dim objFileMgr As New cFileHandler()
        'Dim strNewFileName As String
        Dim strAssemblyName, strAssemblyLocation, strPgNumberFormat As String
        '

        Select Case e.Control.Id
            Case "grpCrypt_getForm"
                'frm_cryptTools = New frm_cryptTools()
                'frm_cryptTools.ShowDialog()
                '
            Case "grpMetaData_Remove_FromDoc"
                'Remove Meta Data on Save
                'https://answers.microsoft.com/en-us/msoffice/forum/all/automatically-remove-metadata-upon-save/8b31c8c9-23e1-490f-8aa7-6bfe661807e5
                'See https://docs.microsoft.com/en-us/office/vba/api/word.wdremovedocinfotype
                'for information on the individual information types if you do not want to remove them all.
                '
                myDoc = objGlobals.glb_get_wrdActiveDoc
                'myDocInfo = New System.IO.FileInfo(myDoc.FullName)
                '
                If Not (myDoc.Path = "") Then
                    If objMsgMgr.msg_warning_MetaDataRemoval() Then
                        '
                        strAssemblyLocation = objPropsMgr.prps_rbn_getAssemblyLocation(myDoc)
                        strAssemblyName = objPropsMgr.prps_rbn_getAssemblyName(myDoc)
                        strPgNumberFormat = objPropsMgr.prps_CustomProperty_get("pgNumberFormat")
                        '
                        objFileMgr.file_get_saveTimeStampedCopy(myDoc, "noMetaData")
                        'strNewFileName = objFileMgr.file_get_newFileName(myDoc, myDocInfo.DirectoryName, "noMetaData")
                        'myDoc.SaveAs2(strNewFileName)
                        '
                        myDoc.RemoveDocumentInformation(WdRemoveDocInfoType.wdRDIAll)
                        '
                        'The following re-establishment routines have been tested and are known
                        'to work
                        Try
                            'objPropsMgr.prps_CustomProperty_set(strAssemblyLocation, "_AssemblyLocation")
                            'objPropsMgr.prps_CustomProperty_set(strAssemblyName, "_AssemblyName")
                        Catch ex As Exception
                            '
                        End Try
                        '
                        '
                        'Re-establish the ribbon as it was in the original document
                        Try
                            'objPropsMgr.prps_CustomProperty_set(strPgNumberFormat, "pgNumberFormat")
                            'objPropsMgr.prps_CustomProperty_set("false", "rptAccessible")
                        Catch ex As Exception

                        End Try

                        '
                        'May need to save the document to make it sticj
                        '
                        'myDocInfo = New System.IO.FileInfo(myDoc.FullName)
                        'myDoc.IsInAutosave
                        myDoc.SaveAs2()
                        '
                    Else

                    End If

                Else
                    MsgBox("Please ensure that your document has been saved before attempting to remove Meta Data.")

                End If
                '
        End Select

    End Sub

    Private Sub PIF_Fin_grpTstLoadFromWeb_Click(sender As Object, e As RibbonControlEventArgs) Handles grpTst_LoadFromWeb_getTemplate.Click, grpTst_LoadFromWeb_getStylesGuide_Accessible.Click, grpTst_LoadFromWeb_getStylesGuide.Click, grpTst_LoadFromResources_getTemplate.Click, grpTst_LoadFromResources_getStylesGuide_Accessible.Click, grpTst_LoadFromResources_getRptPrtExample.Click, grpTst_LoadFromResources_getThemeFile.Click, grpTst_LoadFromResources_getRptLndExample.Click, grpTst_LoadFromResources_getRptBrfExample.Click, grpTst_LoadFromResources_getStylesGuide.Click
        Dim objGlobals As New cGlobals()
        Dim objFileMgr As New cFileHandler()
        Dim objViewMgr As New cViewManager()
        Dim strFilePath, strWebLocation, strSoftwareType As String
        Dim sect As Word.Section
        Dim strMsg As String = ""
        '
        objGlobals.glb_get_wrdApp.ScreenUpdating = False
        '
        strWebLocation = objGlobals.glb_get_webSiteId()
        strSoftwareType = objGlobals.glb_get_SoftwareType()
        '
        sect = objGlobals.glb_get_wrdSect()
        strFilePath = ""
        '
        Try
            Select Case e.Control.Id
                Case "grpTst_LoadFromResources_getTemplate"
                    strFilePath = objFileMgr.file_get_resourcesFromResource("AA_GeneralReport.dotx", "AA_GeneralReport", "")
                    Me.rbn_response_toDownLoad(strFilePath)
                    '
                Case "grpTst_LoadFromResources_getThemeFile"
                    strFilePath = objFileMgr.file_get_resourcesFromResource("AA_ThemeForRpt", "AA_Theme_for_GeneralReport_with_CustClrs_20240808", "")
                    Me.rbn_response_toDownLoad(strFilePath)
                    '
                Case "grpTst_LoadFromResources_getRptPrtExample"
                    strFilePath = objFileMgr.file_get_resourcesFromResource("AA_ReportPrt", "AA_PortraitReport_example", "")
                    Me.rbn_response_toDownLoad(strFilePath)
                    '
                Case "grpTst_LoadFromResources_getRptLndExample"
                    strFilePath = objFileMgr.file_get_resourcesFromResource("AA_ReportLnd", "AA_LandscapeReport_example", "")
                    Me.rbn_response_toDownLoad(strFilePath)
                    '
                Case "grpTst_LoadFromResources_getRptBrfExample"
                    strFilePath = objFileMgr.file_get_resourcesFromResource("AA_ReportBrf", "AA_BriefReport_example", "")
                    Me.rbn_response_toDownLoad(strFilePath)
                    '
                Case "grpTst_LoadFromResources_getStylesGuide"
                    strFilePath = objFileMgr.file_get_resourcesFromResource("AA_StylesGuide", "AA_StylesGuide", "")
                    Me.rbn_response_toDownLoad(strFilePath)
                    objViewMgr.vw_change_toPageFitBestFit(objGlobals.glb_get_wrdActiveDoc())
                    '
                Case "grpTst_LoadFromResources_getStylesGuide_Accessible"
                    strFilePath = objFileMgr.file_get_resourcesFromResource("AA_StylesGuide_Accessible", "AA_StylesGuide_Accessible", "")
                    Me.rbn_response_toDownLoad(strFilePath)
                    objViewMgr.vw_change_toPageFitBestFit(objGlobals.glb_get_wrdActiveDoc())
                    '
                Case "grpTst_LoadFromWeb_getTemplate"
                    strFilePath = objFileMgr.file_get_templateFromWeb(, strWebLocation)
                    If strFilePath = "" Then
                        MsgBox("Error in file download.. Are you connected to the internet?")
                    Else
                        MsgBox("Download complete. The file is located at" + vbCrLf + strFilePath)
                    End If
                    '

                Case "grpTst_LoadFromWeb_getStylesGuide"
                    '
                    'strFilePath = objFileMgr.file_get_resourcesFromWeb("StylesGuide.docx", "exampleDoc", "acilallen.com.au")
                    strFilePath = objFileMgr.file_get_resourcesFromWeb("StylesGuide.docx", "exampleDoc", strWebLocation)
                    Select Case strFilePath
                        Case "cancel"
                            MsgBox("Download has been cancelled by the user")
                        Case ""
                            MsgBox("Error in file download.. Are you connected to the internet?")
                        Case Else
                            strMsg = "Download complete. The file from " + "'" + strWebLocation + "'" + " is located at" + vbCrLf + strFilePath + vbCrLf + vbCrLf + "Do you want the file opened"
                            'MsgBox("Download complete. The file is located at" + vbCrLf + strFilePath)
                            If MsgBox(strMsg, vbYesNo + vbDefaultButton2, "Download status") = MsgBoxResult.Yes Then
                                objGlobals.glb_get_wrdApp.Documents.Open(strFilePath)
                            End If
                    End Select
                    '

                Case "grpTst_LoadFromWeb_getStylesGuide_Accessible"
                    strFilePath = objFileMgr.file_get_resourcesFromWeb("StylesGuide-AccessibleAware.docx", "exampleDoc", strWebLocation)
                    '
                    Select Case strFilePath
                        Case "cancel"
                            MsgBox("Download has been cancelled by the user")
                        Case ""
                            MsgBox("Error in file download.. Are you connected to the internet?")
                        Case Else
                            strMsg = "Download complete. The file from " + "'" + strWebLocation + "'" + " is located at" + vbCrLf + strFilePath + vbCrLf + vbCrLf + "Do you want the file opened"
                            'MsgBox("Download complete. The file is located at" + vbCrLf + strFilePath)
                            If MsgBox(strMsg, vbYesNo + vbDefaultButton2, "Download status") = MsgBoxResult.Yes Then
                                objGlobals.glb_get_wrdApp.Documents.Open(strFilePath)
                            End If
                            '
                    End Select
                    '

            End Select

        Catch ex As Exception

        End Try
        '
        '
        objGlobals.glb_get_wrdApp.ScreenUpdating = True
        '
    End Sub

    Private Sub PIF_PgS_dlg(sender As Object, e As RibbonControlEventArgs) Handles grpRpt_CoversAndTOC.DialogLauncherClick
        Dim obj As New cGetDotNetVersion()
        Dim frm As frm_About
        '
        frm = New frm_About()
        frm.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        frm.Location = System.Windows.Forms.Cursor.Position
        frm.ShowDialog()
        '
    End Sub

    Private Sub PIF_PgS_CoverPageDelete(sender As Object, e As RibbonControlEventArgs) Handles gal_CoverPages.ButtonClick
        Dim objCpMgr As cCoverPageMgr
        Dim myDoc As Word.Document
        Dim sect As Word.Section
        Dim objRptMgr As New cReport()
        Dim objGlb As New cGlobals()
        Dim strMsg As String
        '
        strMsg = ""
        '
        Try
            Select Case e.Control.Id
                Case "gal_CoverPages_btn_deleteCoverPage"
                    myDoc = objGlb.glb_get_wrdActiveDoc
                    sect = Nothing
                    objCpMgr = New cCoverPageMgr()
                    '
                    strMsg = objCpMgr.cp_Delete_CoverPage(myDoc)
                    If strMsg <> "" Then
                        MsgBox(strMsg)
                    End If

            End Select
        Catch ex As Exception

        End Try
        '
    End Sub

    Private Sub PIF_PgS_grpCoveringLetter_Click(sender As Object, e As RibbonControlEventArgs) Handles grpLetter_insertLetter.Click, grpLetter_insertMemo.Click, grpLetter_btn_forMemo.Click, grpLetter_insertMemo_swBuild.Click, grpLetter_insertLetter_swBuild.Click, grpLetter_LtrHead3.Click, grpLetter_LtrHead2.Click, grpLetter_LtrHead1.Click, grpLetter_delReport.Click, grpLetter_Sydney.Click, grpLetter_Perth.Click, grpLetter_Melbourne.Click, grpLetter_Canberra.Click, grpLetter_Brisbane.Click, grpLetter_Adelaide.Click, tbHome_grpLetter_standaloneLetter.Click, tbHome_mnu_contactDetails_letter_Melbourne.Click, tbHome_mnu_contactDetails_letter_Sydney.Click, tbHome_mnu_contactDetails_letter_Brisbane.Click, tbHome_mnu_contactDetails_letter_Adelaide.Click, tbHome_mnu_contactDetails_letter_Perth.Click, tbHome_mnu_contactDetails_letter_Canberra.Click, tbHome_grpLetter_standaloneMemo.Click
        'Dim objSectMgr As New cSectionMgr()
        Dim objGlobals As New cGlobals()
        Dim objToolsMgr As New cTools()
        Dim objMsgMgr As New cMessageManager()
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objPlhMgr As New cPlHBase()
        Dim objChptLetter As New cStationeryLetter()
        Dim objChptMemo As New cStationeryMemo()
        Dim objBBMgr As cBBlocksHandler
        Dim objThmMgr As New cThemeMgr()
        Dim objRptMgr As New cReport()
        Dim objStylesMgr As New cStylesManager()

        Dim objRpt As New cReport()
        Dim myDoc As Word.Document
        Dim sect As Word.Section
        Dim rng As Range
        Dim tbl As Word.Table
        Dim drCell As Word.Cell
        Dim rslt As Boolean
        Dim para As Word.Paragraph
        Dim strMsg As String
        '
        'objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = False
        objGlobals.glb_screen_update(False)
        '
        strMsg = ""
        '
        Select Case e.Control.Id
            Case "tbHome_grpLetter_standaloneLetter"
                myDoc = objGlobals.glb_get_wrdApp.Documents.Add()
                myDoc.AttachedTemplate = objGlobals.glb_getTmpl_FullName()
                objThmMgr.thm_Set_ThemeToAAStd_fromFile(objGlobals.glb_get_wrdActiveDoc)
                objRptMgr.Rpt_Styles_resetStyles_fromTemplate(True)
                '
                'objGlobals.glb_doc_checkDocType_ActivateTab(objGlobals._strTabId_ThemesAndTest)

                'myDoc.AttachedTemplate = objGlobals.glb_getTmpl_FullName()
                '
                rng = myDoc.Sections.First.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Select()
                '
                'Create instance of objBBMgr here so that we pick up the correct template
                'If created too early it will set the template to an earlier attached template
                objBBMgr = New cBBlocksHandler()
                rng = objBBMgr.insertBuildingBlockFromDefaultLib_ReturnRange("aa_letter", "aa_Stationery")
                rng.Select()
                '
                objRpt.Rpt_delete_ReportSections()
                '
                objStylesMgr.glb_doc_checkDocType_ActivateTab(objGlobals._strTabId_PagesAndSections)

            Case "grpLetter_insertLetter"
                rng = objGlobals.glb_get_wrdActiveDoc().Sections.First.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Select()
                '
                objBBMgr = New cBBlocksHandler()
                rng = objBBMgr.insertBuildingBlockFromDefaultLib_ReturnRange("aa_letter", "aa_Stationery")
                rng.Select()
                '
            Case "grpLetter_insertLetter_swBuild"
                rng = objChptLetter.ltr_insert_Letter_Memo()
                rng.Select()
                '
            Case "tbHome_grpLetter_standaloneMemo"
                myDoc = objGlobals.glb_get_wrdApp.Documents.Add()
                myDoc.AttachedTemplate = objGlobals.glb_getTmpl_FullName()
                objThmMgr.thm_Set_ThemeToAAStd_fromFile(objGlobals.glb_get_wrdActiveDoc)
                objRptMgr.Rpt_Styles_resetStyles_fromTemplate(True)
                '
                'objGlobals.glb_doc_checkDocType_ActivateTab(objGlobals._strTabId_ThemesAndTest)

                'myDoc.AttachedTemplate = objGlobals.glb_getTmpl_FullName()
                '
                rng = myDoc.Sections.First.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Select()
                '
                'Create instance of objBBMgr here so that we pick up the correct template
                'If created too early it will set the template to an earlier attached template
                objBBMgr = New cBBlocksHandler()
                rng = objBBMgr.insertBuildingBlockFromDefaultLib_ReturnRange("aa_memo", "aa_Stationery")
                rng.Select()
                '
                objRpt.Rpt_delete_ReportSections()
                '
                objStylesMgr.glb_doc_checkDocType_ActivateTab(objGlobals._strTabId_PagesAndSections)

            Case "grpLetter_insertMemo"
                rng = objGlobals.glb_get_wrdActiveDoc().Sections.First.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Select()
                '
                objBBMgr = New cBBlocksHandler()
                rng = objBBMgr.insertBuildingBlockFromDefaultLib_ReturnRange("aa_memo", "aa_Stationery")
                sect = rng.Sections.First
                rng = sect.Range
                '
                If rng.Tables.Count <> 0 Then
                    tbl = rng.Tables.Item(1)
                    drCell = tbl.Range.Cells.Item(3)
                    rng = drCell.Range
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    Call rng.MoveEnd(WdUnits.wdWord, 1)
                    rng.Select()
                End If
                '
            Case "grpLetter_insertMemo_swBuild"
                '
                'rng = objChptMemo.Insert_Stationery_Memo()
                rng = objChptLetter.ltr_insert_Letter_Memo("memo")
                '
                If rng.Tables.Count <> 0 Then
                    tbl = rng.Tables.Item(1)
                    drCell = tbl.Range.Cells.Item(3)
                    rng = drCell.Range
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    Call rng.MoveEnd(WdUnits.wdWord, 1)
                    rng.Select()
                End If
                '
                '
            Case "grpLetter_insertBrief"
                rng = objChptLetter.ltr_insert_Letter_Memo("briefingNote")
                '
                If rng.Tables.Count <> 0 Then
                    tbl = rng.Tables.Item(1)
                    drCell = tbl.Range.Cells.Item(3)
                    rng = drCell.Range
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    Call rng.MoveEnd(WdUnits.wdWord, 1)
                    rng.Select()
                End If
                '
            Case "grpLetter_Melbourne", "tbHome_mnu_contactDetails_letter_Melbourne"
                objChptLetter.ltr_insert_OfficeAddress_intoFooter("Melbourne")
                '
            Case "grpLetter_Sydney", "tbHome_mnu_contactDetails_letter_Sydney"
                objChptLetter.ltr_insert_OfficeAddress_intoFooter("Sydney")
                '
            Case "grpLetter_Brisbane", "tbHome_mnu_contactDetails_letter_Brisbane"
                objChptLetter.ltr_insert_OfficeAddress_intoFooter("Brisbane")
                '
            Case "grpLetter_Canberra", "tbHome_mnu_contactDetails_letter_Canberra"
                objChptLetter.ltr_insert_OfficeAddress_intoFooter("Canberra")
                '
            Case "grpLetter_Perth", "tbHome_mnu_contactDetails_letter_Perth"
                objChptLetter.ltr_insert_OfficeAddress_intoFooter("Perth")
                '
            Case "grpLetter_Hobart", "tbHome_mnu_contactDetails_letter_Hobart"
                objChptLetter.ltr_insert_OfficeAddress_intoFooter("Hobart")
                '
            Case "grpLetter_Adelaide", "tbHome_mnu_contactDetails_letter_Adelaide"
                objChptLetter.ltr_insert_OfficeAddress_intoFooter("Adelaide")
                '
            Case "grpLetter_Memo"
                MsgBox("This function is no longer necessary, so it has been discontinued")
                'Call objSectMgr.objHfMgr.hf_footers_LetterContacts_Clear(objSectMgr.currentSect)
                'rng = objSectMgr.currentSect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Range
                'rng.Collapse(WdCollapseDirection.wdCollapseStart)
                'Call objSectMgr.objBBMgr.insertBuildingBlockFromDefaultLibToRange("txtBx_contact_Memo", "letters", rng)
            Case "grpLetter_BriefingNote"
                MsgBox("This function is no longer necessary, so it has been discontinued")
                'Call objSectMgr.objHfMgr.hf_footers_LetterContacts_Clear(objSectMgr.currentSect)
                'rng = objSectMgr.currentSect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Range
                'rng.Collapse(WdCollapseDirection.wdCollapseStart)
                'Call objSectMgr.objBBMgr.insertBuildingBlockFromDefaultLibToRange("txtBx_contact_Memo", "letters", rng)
            Case "grpLetter_LtrHead1"
                objGlobals.glb_get_wrdApp.Selection.Style = objGlobals.glb_get_wrdApp.ActiveDocument.Styles("Letter Heading 1")
                para = objGlobals.glb_get_wrdApp.Selection.Paragraphs.Item(1)
                para.Format.KeepWithNext = True
            Case "grpLetter_LtrHead2"
                objGlobals.glb_get_wrdApp.Selection.Style = objGlobals.glb_get_wrdApp.ActiveDocument.Styles("Letter Heading 2")
                para = objGlobals.glb_get_wrdApp.Selection.Paragraphs.Item(1)
                para.Format.KeepWithNext = True
            Case "grpLetter_LtrHead3"
                objGlobals.glb_get_wrdApp.Selection.Style = objGlobals.glb_get_wrdApp.ActiveDocument.Styles("Letter Heading 3")
                para = objGlobals.glb_get_wrdApp.Selection.Paragraphs.Item(1)
                para.Format.KeepWithNext = True
            Case "grpLetter_delReport"
                rslt = objMsgMgr.deleteReportMessage
                If rslt Then
                    '
                    objRpt.Rpt_delete_ReportSections()
                    '
                    '
                    '
                Else
                    'MsgBox ("No was chosen")
                End If
            Case "grpLetter_convertCaption"
                rng = objGlobals.glb_get_wrdApp.Selection.Range
                'Call objPlhMgr.convertBoxesTo("letter")
                rng.Select()
                'Call objPlhMgr.convertTablesTo("letter")
                rng.Select()
                'Call objPlhMgr.convertFiguresTo("letter")
                rng.Select()
            Case "grpLetter_convertCaptionBoxes"
                rng = objGlobals.glb_get_wrdApp.Selection.Range
                'Call objPlhMgr.convertBoxesTo("letter")
                rng.Select()
            Case "grpLetter_convertCaptionFigures"
                rng = objGlobals.glb_get_wrdApp.Selection.Range
                'Call objPlhMgr.convertFiguresTo("letter")
                rng.Select()
            Case "grpLetter_convertCaptionTables"
                rng = objGlobals.glb_get_wrdApp.Selection.Range
                'Call objPlhMgr.convertTablesTo("letter")
                rng.Select()
            Case "grpLetter_convertCaptionKeyFinding"
                rng = objGlobals.glb_get_wrdApp.Selection.Range
                'Call objPlhMgr.convertBoxesKeyFindingTo("base")
                rng.Select()
            Case "grpLetter_convertCaptionRecommendation"
                rng = objGlobals.glb_get_wrdApp.Selection.Range
                'Call objPlhMgr.convertBoxesRecommendationsTo("base")
                rng.Select()
            Case Else
        End Select
        '
        'objSectMgr.objGlobals.glb_get_wrdApp.ScreenUpdating = True
        objGlobals.glb_screen_update(True)

        'Re-establish user screen updating option
        '
    End Sub

    Private Sub PIF_PgS_grpWhatsNew_Click(sender As Object, e As RibbonControlEventArgs) Handles grpWhatsNew_Form.Click
        Dim objGlobals As New cGlobals
        Dim frm As New frm_WhatsNew()
        '
        'objGlobals.glb_screen_update()
        '
        Select Case e.Control.Id
            Case "grpWhatsNew_Form"
                frm.Show()
        End Select

    End Sub

    Private Sub btn_colorPicker_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_colorPicker.Click
        Dim frm As New frm_colorPicker("text_Colour")
        '
        frm.Show()
    End Sub

    Private Sub btn_ApplyStdTheme_Manually_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_ApplyStdTheme_Manually.Click
        Dim objThm As New cThemeMgr()
        '
        objThm.thm_Set_ThemeToAAStd_Manually(Globals.ThisAddIn.Application.ActiveDocument)
        MsgBox("Theme applied")
        '
    End Sub

    Private Sub btn_Crypt_Click(sender As Object, e As RibbonControlEventArgs)
        'Dim frm_Crypt As New frm_cryptTools()
        '
        'frm_Crypt.Show()
    End Sub

    Private Sub btn_styles_makeTableText_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_styles_makeTableText.Click
        Dim objGlobals As New cGlobals()
        Dim objStylesMgr As New cStylesManager()
        Dim objTblStyles As New cTableStyles()
        Dim myStyle As Word.Style
        '
        'var_tbl_myDoc = glb_get_wrdActiveDoc()
        '
        'Me.var_tbl_colourHeader = Me._glb_colour_purple_Dark
        'Me.var_tbl_colourUnits = Me._glb_colour_UnitsGrey
        '
        'myStyle = objStylesMgr.style_getCreateRefresh_Table_text(objGlobals.glb_get_wrdActiveDoc)
        myStyle = objStylesMgr.style_txt_getTableTextStyle(objGlobals.glb_get_wrdActiveDoc)
        myStyle = objStylesMgr.style_txt_getTableHeadingStyle(objGlobals.glb_get_wrdActiveDoc)
        myStyle = objStylesMgr.style_txt_getTableUnitsRowStyle(objGlobals.glb_get_wrdActiveDoc)
        '
        objTblStyles.tblstyl_add_aacTableBasic(objGlobals.glb_get_wrdActiveDoc)
        objTblStyles.tblstyl_add_aacTableNoLines(objGlobals.glb_get_wrdActiveDoc)
        '
    End Sub

    Private Sub tbHome_btn_Help_Click(sender As Object, e As RibbonControlEventArgs) Handles tbHome_btn_Help.Click
        MsgBox("Help information in the form of documents and/or videos can be made available under this button")
    End Sub
End Class
