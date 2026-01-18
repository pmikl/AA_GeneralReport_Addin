Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports System.Drawing
'
'rev 01.00  20250830
'
Public Class cImageGetAndEdit
    '
    Public objGlobals As cGlobals
    Public objCropMgr As cCropMgr
    '
    Public Sub New()
        MyBase.New()
        Me.objGlobals = New cGlobals()
        Me.objCropMgr = New cCropMgr()
    End Sub
    '
    Public Sub imgGet_fill_backPanelFromFile_aac_BackColour()
        Dim strMsg As String

        strMsg = ""

        strMsg = Me.imgGet_fill_backPanelFromFile("aac_BackColour")
        '
        Select Case strMsg
            Case "ok"
            Case "cropFailure"
                MsgBox("Software fault. No new cropped image")
            Case "cropCancel"
                MsgBox("Cropping function cancelled by the user")
            Case "no panel"
                MsgBox("The current section does not have an image back panel")
            Case Else
        End Select

    End Sub
    '
    '
    Public Sub imgGet_fill_backPanelFromFile_cp_pict_large()
        Dim strMsg As String

        strMsg = ""

        strMsg = Me.imgGet_fill_backPanelFromFile("cp_pict_large")
        '
        Select Case strMsg
            Case "ok"
            Case "cropFailure"
                MsgBox("Cropping failure. No new cropped image")
            Case "cropCancel"
                MsgBox("Cropping function cancelled by the user")
            Case "no panel"
                MsgBox("The current section does not have an small Cover Page picture panel")
            Case Else
        End Select

    End Sub
    '
    '
    ''' <summary>
    ''' Will fill the back panel in the current section with a cropped image from file. The 
    ''' Back Panel is identified by name. At present 'aac_BackColour' for the legacy purple
    ''' back panels and 'cp_pict_large' for the small Cover Page Panel
    ''' </summary>
    ''' <param name="strBackPanelName"></param>
    Public Function imgGet_fill_backPanelFromFile(Optional strBackPanelName As String = "aac_BackColour") As String
        Dim objPanelMgr As New cBackPanelMgr()
        Dim lstOfBackPanels As List(Of cShapeMgr)
        Dim objBackPanel As New cShapeMgr()
        Dim objViewMgr As New cViewManager()
        '
        Dim frm As frm_pictControl3
        Dim objSrcImg, cropRect As Word.Shape
        Dim dlg_Picture As Word.Dialog
        Dim strFilePath As String
        Dim sect As Word.Section
        Dim sel As Word.Selection
        '
        Dim strMsg As String
        '
        strFilePath = ""
        sect = Me.objGlobals.glb_get_wrdSect()
        '
        strMsg = ""
        '
        If objPanelMgr.pnl_has_BackPanel(sect, strBackPanelName) Then
            'sect = Globals.ThisDocument.Application.Selection.Sections.Item(1)
            lstOfBackPanels = objPanelMgr.pnl_getBackPanel_PlaceHolders(sect, strBackPanelName)
            '
            dlg_Picture = Me.objGlobals.glb_get_wrdApp.Dialogs(Word.WdWordDialog.wdDialogInsertPicture)
            dlg_Picture.name = "*.*"
            strFilePath = ""
            '
            If dlg_Picture.Display() = -1 Then
                strFilePath = dlg_Picture.Name
                Try
                    If lstOfBackPanels.Count > 0 Then
                        objViewMgr.vw_change_ColumnsAndRows(sect)
                        '
                        objBackPanel = lstOfBackPanels.Item(0)
                        '
                        'Place the Source Image
                        Me.objGlobals.glb_get_wrdApp.ScreenUpdating = False
                        '
                        objCropMgr.crp_setView_toHeader()
                        'Need to change the next line to accept image from file
                        objSrcImg = objCropMgr.crp_SrcImage_Place(objBackPanel, sect, "file", strFilePath)       'This is the image to be cropped
                        '
                        'Calculate aspect ration of objSrcImg.. If within tolerance don't go through cropping, just insert
                        'direct from file, after deleting the placed image of course.
                        '
                        Dim pnlAspectRatio, srcAspectRatio As Single
                        pnlAspectRatio = objBackPanel.aspectRatio                           'verified correct for A4 1.414
                        srcAspectRatio = objSrcImg.Height / objSrcImg.Width                 'calculated as 1, verified as 1
                        '
                        '***
                        '
                        'MsgBox("Image ratio h/w = " + srcAspectRatio.ToString() + vbCrLf + "Panel ratio h/w = " + pnlAspectRatio.ToString())
                        '
                        If sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait Then
                            If (objSrcImg.Height * (1 - 0.1)) / ((1 + 0.1) * objSrcImg.Width) <= objBackPanel.aspectRatio And objBackPanel.aspectRatio <= (objSrcImg.Height * (1 + 0.1)) / ((1 - 0.1) * objSrcImg.Width) Then
                                'Image aspect ratio is withing tolerance so we
                                'do raw image fill
                                'MsgBox("OK for raw image fill")

                            Else
                                'Do the normal thing



                            End If
                        End If
                        '
                        '*** ???
                        If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then
                            If (objSrcImg.Height * (1 - 0.1)) / ((1 + 0.1) * objSrcImg.Width) <= objBackPanel.aspectRatio And objBackPanel.aspectRatio <= (objSrcImg.Height * (1 + 0.1)) / ((1 - 0.1) * objSrcImg.Width) Then
                                'Image aspect ratio is withing tolerance so we
                                'do raw image fill
                                'MsgBox("OK for raw image fill")

                            Else
                                'Do the normal thing



                            End If
                        End If
                        '

                        '
                        '***
                        '
                        cropRect = objCropMgr.crp_insert_CropRect(sect, objBackPanel, objSrcImg)
                            objCropMgr.crp_setView_toPrintView()
                            '
                            objGlobals.glb_screen_update(True)
                            '
                            frm = New frm_pictControl3(objCropMgr, objSrcImg, objBackPanel)
                            '
                            'OK = 1, Cancel = 2
                            If frm.ShowDialog() = 1 Then
                                '
                                'objSrcImg is now cropped
                                '
                                objCropMgr.crp_setView_toHeader()
                                objSrcImg.Select()
                                '
                                If Not IsNothing(objCropMgr.objShp_NewPic.shp) Then
                                    '
                                    objCropMgr.crp_setView_toHeader()
                                    objCropMgr.objShp_NewPic.shp.Select()
                                    sel = objGlobals.glb_get_wrdApp.Selection
                                    sel.CopyAsPicture()
                                    objCropMgr.crp_delete_CropRect_and_SrcImage()
                                    '
                                    objCropMgr.crp_fill_ShapeWithImage()
                                    '
                                    objCropMgr.crp_setView_toPrintView()
                                    'objViewMgr.vw_change_ColumnsAndRows(sect, 2)
                                    strMsg = "ok"
                                Else
                                    strMsg = "cropFailure"
                                End If
                            Else
                                objCropMgr.crp_delete_CropRect_and_SrcImage()
                                strMsg = "cropCancel"
                            End If

                        Else
                            strMsg = "no panel"
                    End If

                Catch ex As Exception
                    '
                    objGlobals.glb_screen_update(True)
                    '
                End Try

            Else
                'Cancelled by the user
            End If
        Else
            strMsg = "no panel"
        End If
        '
        Return strMsg
        '
    End Function
    '
    Public Sub imgGet_fill_withRawImageFromFile_aac_BackColour()
        Dim strMSg As String

        strMSg = imgGet_fill_withRawImageFromFile("aac_BackColour")

        Select Case strMsg
            Case "ok"
            Case "error"
                MsgBox("Software Try/Catch error in objPanelMgr.pnl_fill_withRawUserImage")
            Case "no panel"
                MsgBox("The current section does not have an image back panel")
            Case Else
        End Select

    End Sub
    '
    Public Sub imgGet_fill_withRawImageFromFile_cp_pict_large()
        Dim strMSg As String

        strMSg = imgGet_fill_withRawImageFromFile("cp_pict_large")

        Select Case strMSg
            Case "ok"
            Case "error"
                MsgBox("Software Try/Catch error in objPanelMgr.pnl_fill_withRawUserImage")
            Case "no panel"
                MsgBox("The current section does not have an small Cover Page picture panel")
            Case Else
        End Select

    End Sub


    Public Function imgGet_fill_withRawImageFromFile(Optional strBackPanelName As String = "aac_BackColour") As String
        Dim objPanelMgr As New cBackPanelMgr()
        Dim lstOfBackPanels As List(Of cShapeMgr)
        Dim objShpMgr As cShapeMgr
        Dim sect As Word.Section
        Dim strMSg As String
        '
        strMSg = ""
        '
        sect = objGlobals.glb_get_wrdSect()

        If objPanelMgr.pnl_has_BackPanel(sect, strBackPanelName) Then
            '
            lstOfBackPanels = objPanelMgr.pnl_getBackPanel_PlaceHolders(sect, strBackPanelName)
            objShpMgr = lstOfBackPanels.Item(0)
            '
            Try
                'Globals.ThisDocument.Application.ScreenRefresh()
                objPanelMgr.pnl_fill_withRawUserImage(objShpMgr)
                '
                strMSg = "ok"
                '
            Catch ex As Exception
                strMSg = "error"
            End Try
            '
            'objGlobals.glb_cursors_setToNormal()
            '
        Else
            strMSg = "no panel"
        End If
        '
        Return strMSg

    End Function
    '
    '
    Public Sub imgGet_fill_backPanelFromClipboard_aac_BackColour()
        Dim strMsg As String
        Dim objMsgMgr As New cMessageManager()
        '
        strMsg = Me.imgGet_fill_backPanelFromClipboard("aac_BackColour")
        '
        Select Case strMsg
            Case "cropFailure"
                MsgBox("Cropping failure. No new cropped image")
            Case "cropCancel"
                MsgBox("Cropping function cancelled by the user")
            Case "error01"
                MsgBox("General Try/Catch error")
            Case "emptyClipboard"
                objMsgMgr.msg_cropping_Error()
            Case "no panel"
                MsgBox("The current section does not have an image back panel")
            Case Else

        End Select
    End Sub
    '
    Public Sub imgGet_fill_backPanelFromClipboard_cp_pict_large()
        Dim strMsg As String
        Dim objMsgMgr As New cMessageManager()
        '
        strMsg = Me.imgGet_fill_backPanelFromClipboard("cp_pict_large")
        '
        Select Case strMsg
            Case "cropFailure"
                MsgBox("Cropping failure. No new cropped image")
            Case "cropCancel"
                MsgBox("Cropping function cancelled by the user")
            Case "error01"
                MsgBox("General Try/Catch error")
            Case "emptyClipboard"
                objMsgMgr.msg_cropping_Error()
            Case "no panel"
                MsgBox("The current section does not have an image back panel")
            Case Else

        End Select
    End Sub



    '
    Public Function imgGet_fill_backPanelFromClipboard(Optional strBackPanelName As String = "aac_BackColour") As String
        Dim objPanelMgr As New cBackPanelMgr()
        Dim lstOfBackPanels As List(Of cShapeMgr)
        Dim objBackPanel As New cShapeMgr()
        Dim objViewMgr As New cViewManager()
        '
        Dim frm As frm_pictControl3
        Dim objSrcImg, cropRect As Word.Shape
        Dim strFilePath As String
        Dim sect As Word.Section
        Dim sel As Word.Selection
        Dim strMsg As String
        '
        strFilePath = ""
        strMsg = ""
        sect = Me.objGlobals.glb_get_wrdSect()


        If objPanelMgr.pnl_has_BackPanel(sect) Then
            'sect = Globals.ThisDocument.Application.Selection.Sections.Item(1)
            lstOfBackPanels = objPanelMgr.pnl_getBackPanel_PlaceHolders(sect, strBackPanelName)
            '
            If System.Windows.Forms.Clipboard.ContainsImage() Then
                Try
                    If lstOfBackPanels.Count > 0 Then
                        objViewMgr.vw_change_ColumnsAndRows(sect)
                        '
                        objBackPanel = lstOfBackPanels.Item(0)
                        '
                        '
                        'Place the Source Image
                        Me.objGlobals.glb_get_wrdApp.ScreenUpdating = False
                        'Globals.ThisDocument.Application.ScreenUpdating = False
                        '
                        objCropMgr.crp_setView_toHeader()
                        objSrcImg = objCropMgr.crp_SrcImage_Place(objBackPanel, sect)       'This is the image to be cropped
                        cropRect = objCropMgr.crp_insert_CropRect(sect, objBackPanel, objSrcImg)
                        objCropMgr.crp_setView_toPrintView()
                        '
                        Me.objGlobals.glb_get_wrdApp.ScreenUpdating = True
                        'Globals.ThisDocument.Application.ScreenUpdating = True
                        '
                        frm = New frm_pictControl3(objCropMgr, objSrcImg, objBackPanel)
                        '
                        'OK = 1, Cancel = 2
                        If frm.ShowDialog() = 1 Then
                            '
                            'objSrcImg is now cropped
                            '
                            objCropMgr.crp_setView_toHeader()
                            objSrcImg.Select()
                            '
                            If Not IsNothing(objCropMgr.objShp_NewPic.shp) Then
                                '
                                objCropMgr.crp_setView_toHeader()
                                objCropMgr.objShp_NewPic.shp.Select()
                                sel = Me.objGlobals.glb_get_wrdApp.Selection
                                sel.CopyAsPicture()
                                objCropMgr.crp_delete_CropRect_and_SrcImage()
                                '
                                objCropMgr.crp_fill_ShapeWithImage()
                                '
                                objCropMgr.crp_setView_toPrintView()
                                'objViewMgr.vw_change_ColumnsAndRows(sect, 2)
                            Else
                                strMsg = "cropFailure"
                            End If
                        Else
                            objCropMgr.crp_delete_CropRect_and_SrcImage()
                            strMsg = "cropCancel"
                        End If

                    Else
                        strMsg = "no panel"
                    End If

                Catch ex As Exception
                    '
                    strMsg = "error01"
                    objGlobals.glb_screen_update(True)
                    '
                End Try
            Else
                strMsg = "emptyClipboard"
            End If

        Else
            strMsg = "no panel"
        End If
        '
        Return strMsg
        '
    End Function

End Class
