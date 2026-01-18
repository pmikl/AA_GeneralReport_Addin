Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cImageMgr

    Public currentSect As Section
    Public newShape As Word.Shape
    Public strPicturePlaceHolderName As String
    Public strFormLayoutName As String
    Public objTools As cTools
    '
    Public Sub New()
        Me.objTools = New cTools()
        Me.strFormLayoutName = "Acil Allen"
        Me.strPicturePlaceHolderName = "Logo_Pict_Background"
        '
    End Sub
    '
    Public Function img_get_ImageAsShape(ByRef drCell As Word.Cell) As Word.Shape
        Dim newPic As Word.Shape
        Dim rng As Word.Range
        '
        rng = drCell.Range
        '
        newPic = Nothing
        Try
            If rng.InlineShapes.Count > 0 Then
                newPic = rng.InlineShapes.Item(rng.InlineShapes.Count).ConvertToShape()
                newPic.Select()
            Else
                If rng.ShapeRange.Count > 0 Then
                    newPic = rng.ShapeRange.Item(1)
                End If
            End If

        Catch ex As Exception

        End Try
        '
        Return newPic
    End Function

    '
    Public Function img_get_ImageAsShape(ByRef rng As Word.Range) As Word.Shape
        Dim newPic As Word.Shape
        '
        newPic = Nothing
        '
        If rng.InlineShapes.Count > 0 Then
            newPic = rng.InlineShapes.Item(rng.InlineShapes.Count).ConvertToShape()
            'newPic = rng.InlineShapes.Item(1).ConvertToShape()
            newPic.Select()
            'Globals.ThisAddin.Application.Selection.Cut()
            'rng = para.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            'rng.Paste()
            '
            'Globals.ThisAddin.Application.ScreenRefresh()
            '
            'newPic = rng.ShapeRange.Item(1)
            'i = 1
        Else
            If rng.ShapeRange.Count > 0 Then
                'Globals.ThisAddin.Application.Selection.Cut()
                '
                'rng = para.Range
                'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                'rng.Paste()
                'Globals.ThisAddin.Application.ScreenRefresh()
                '
                newPic = rng.ShapeRange.Item(1)

            End If
        End If
        '
        Return newPic
    End Function
    '
    '
    ''' <summary>
    ''' This method will first look for inLine SHapes in the range (rng). If it finds any it will
    ''' return the first inLineSHape in the Collection. If it doesn't it looks for shapes in the
    ''' ShapeRange. If it finds any it will convert the first into an inLine Shape and return it
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <returns></returns>
    Public Function img_get_ImageAsInlineShape(ByRef rng As Word.Range) As Word.InlineShape
        Dim shp As Word.Shape
        Dim iShp As Word.InlineShape
        '
        iShp = Nothing
        '
        If rng.InlineShapes.Count > 0 Then
            iShp = rng.InlineShapes.Item(1)
        Else
            If rng.ShapeRange.Count > 0 Then
                shp = rng.ShapeRange.Item(1)
                iShp = shp.ConvertToInlineShape()
                '
            End If
        End If
        '
        Return iShp
    End Function
    '
    '
    ''' <summary>
    ''' This method will first look for inLine SHapes in the range (rng). If it finds any it will
    ''' return the first inLineSHape in the Collection. If it doesn't it looks for shapes in the
    ''' ShapeRange. If it finds any it will convert the first into an inLine Shape and return it
    ''' </summary>
    ''' <param name="drCell"></param>
    ''' <returns></returns>
    Public Function img_get_ImageAsInlineShape(ByRef drCell As Word.Cell) As Word.InlineShape
        Dim shp As Word.Shape
        Dim iShp As Word.InlineShape
        Dim rng As Word.Range
        '
        iShp = Nothing
        rng = drCell.Range()
        '
        If rng.InlineShapes.Count > 0 Then
            iShp = rng.InlineShapes.Item(1)
        Else
            If rng.ShapeRange.Count > 0 Then
                shp = rng.ShapeRange.Item(1)
                iShp = shp.ConvertToInlineShape()
                '
            End If
        End If
        '
        Return iShp
    End Function
    '

    '
    Public Function reinsertPicturePlaceHolder(strPlaceHolderName As String, ByRef rng As Word.Range) As Word.Shape
        'This function will reinsert the PicturePlaceHolder.. Generally called under error conditions
        'such as an frm_pictControl.Cancel or if there was an attempted paste when the clipboard was empty

        Dim objBBMgr As cBBlocksHandler
        Dim rngOfShape As Word.Range
        '
        objBBMgr = New cBBlocksHandler
        rngOfShape = objBBMgr.insertBuildingBlockFromDefaultLibToRange(strPlaceHolderName, "CoverPage", rng)
        If rngOfShape.ShapeRange.Count <> 0 Then
            Me.newShape = rngOfShape.ShapeRange.Item(1)
            'Call shp.ZOrder(msoSendBehindText)
            'Call shp.ZOrder(msoSendToBack)
        End If
        reinsertPicturePlaceHolder = Me.newShape
    End Function
    '
    Public Function getPicturePlaceHolders(ByRef sect As Word.Section) As Collection
        'This method returns a collection of the picture placeholders
        'in the section.. They are isdentified by name adn returned as
        'cShapeMgr types.. The original shapes are deleted because we have all their
        'details
        '
        Dim hf As Word.HeaderFooter
        Dim objShape As cShapeMgr
        Dim shp As Word.Shape
        '
        getPicturePlaceHolders = New Collection()
        hf = sect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        '
        For Each shp In hf.Shapes
            If shp.Name Like "cp_pict_*" Then
                objShape = New cShapeMgr
                Call objShape.InitShape(shp, hf)
                Call getPicturePlaceHolders.Add(objShape)
            End If
        Next shp
        '
        For Each shp In hf.Shapes
            If shp.Name Like "cp_pict_*" Then
                shp.Delete()
            End If
        Next shp
        '
    End Function
    '
    ''' <summary>
    ''' This method will return any back panel shapes in the section as a collection of
    ''' cShapeMgr. The backpanel is selected by name "aac_BackColour". Existing back panels are deleted
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function getBackPanel_PlaceHolders(ByRef sect As Word.Section) As List(Of cShapeMgr)
        'This method returns a collection of the picture placeholders
        'in the section.. They are isdentified by name adn returned as
        'cShapeMgr types.. The original shapes are deleted because we have all their
        'details
        '
        Dim hf As Word.HeaderFooter
        Dim objShape As cShapeMgr
        Dim i As Int16
        Dim shp As Word.Shape
        Dim secttest_00 As Word.Section
        Dim lstOfShapes As New List(Of cShapeMgr)
        Dim strShapeName = "aac_BackColour"
        '
        hf = Nothing
        Me.currentSect = sect
        objShape = Nothing
        '
        For Each hf In sect.Headers
            If hf.Exists And hf.IsHeader Then
                If hf.Index = WdHeaderFooterIndex.wdHeaderFooterFirstPage Or hf.Index = WdHeaderFooterIndex.wdHeaderFooterPrimary Then
                    For Each shp In hf.Range.ShapeRange
                        If shp.Name = strShapeName Then
                            objShape = New cShapeMgr
                            objShape.InitShape(shp, hf)
                            lstOfShapes.Add(objShape)
                            shp.Delete()
                            GoTo loop00
                        End If
                    Next shp
                End If
            End If
        Next
        '
loop00:
        If lstOfShapes.Count > 0 Then
            secttest_00 = objShape.hf.Range.Sections.Item(1)
            i = 0
        End If
        '
        GoTo loop_01
        '
        If sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Exists Then
            hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        Else
            hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        End If
        '
        For Each shp In hf.Shapes
            If shp.Name = strShapeName Then
                objShape = New cShapeMgr
                objShape.InitShape(shp, hf)
                lstOfShapes.Add(objShape)
            End If
        Next shp
        '
        For Each shp In hf.Shapes
            If shp.Name = strShapeName Then
                shp.Delete()
            End If
        Next shp
        '
loop_01:

        Return lstOfShapes
    End Function
    '
    '
    ''' <summary>
    ''' Version built 20231022
    ''' </summary>
    ''' <param name="InsertOrPaste"></param>
    ''' <param name="listOfShapes"></param>
    ''' <returns></returns>
    Public Function BackPanel_Replace(InsertOrPaste As String, ByRef listOfShapes As List(Of cShapeMgr)) As Word.Shape
        Dim picDialog As Word.Dialog
        Dim objShp As cShapeMgr
        Dim sect As Word.Section
        Dim hf As Word.HeaderFooter
        Dim rng As Word.Range
        Dim NewPic As Word.Shape
        Dim cropRect As Word.Shape
        Dim frm As frm_pictControl2
        '
        objShp = listOfShapes.Item(0)
        sect = objShp.anchor.Sections.Item(1)
        Me.currentSect = sect
        '
        rng = objShp.anchor
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        BackPanel_Replace = Nothing
        NewPic = Nothing
        '
        Try
            picDialog = Globals.ThisAddin.Application.Dialogs(Word.WdWordDialog.wdDialogInsertPicture)
            Me.strPicturePlaceHolderName = objShp.name
            hf = objShp.hf
            '
            Select Case InsertOrPaste
                Case "pasteImage"
                    'Call hf.rng.PasteSpecial(, , wdFloatOverText, , wdPasteMetafilePicture)
                    Call rng.PasteSpecial(, , WdOLEPlacement.wdFloatOverText, , WdPasteDataType.wdPasteMetafilePicture)
                    'HFRange.PasteSpecial Placement:=wdFloatOverText, DataType:=wdPasteMetafilePicture
                    NewPic = rng.ShapeRange(rng.ShapeRange.Count)
                    'Set NewPic = HFSHapes.Item(HFSHapes.Count)
                    NewPic.Top = Me.objTools.tools_math_MillimetersToPoints(297.0#) - NewPic.Height - 40.0#
                    NewPic.Name = objShp.name
        '
                Case "insertImage"
                    picDialog.name = "*.*"
                    If picDialog.Display = 0 Then
                        Globals.ThisAddin.Application.ActiveDocument.Undo(1)
                        Exit Function
                    End If
                    If picDialog.name = "" Then
                        Globals.ThisAddin.Application.ActiveDocument.Undo(1)
                        Exit Function
                    End If
                    '
                    NewPic = hf.Shapes.AddPicture(picDialog.name, False, True, 0, 0, , , rng)
                    NewPic.Top = Me.objTools.tools_math_MillimetersToPoints(297.0#) - NewPic.Height - 40.0#
                    'NewPic.name = strShapeName
                    NewPic.Name = objShp.name
            End Select
            '
            '
            Call NewPic.ScaleWidth(1.0#, MsoTriState.msoTrue)  'sort of resets the record of width & height?
            Call NewPic.ScaleHeight(1.0#, MsoTriState.msoTrue)
            '
            objShp.width_original = NewPic.Width
            objShp.height_original = NewPic.Height

            Call Me.setImagePageProperties(NewPic)          'Page positional properties
            Call Me.setImageScale(NewPic, sect)
            objShp.scaleFactor_W = objShp.width_original / NewPic.Width
            Call Me.setImagePosition(NewPic, sect)
            '
            BackPanel_Replace = NewPic
            '
            If Globals.ThisAddin.Application.ActiveDocument.ProtectionType <> WdProtectionType.wdNoProtection Then Call Globals.ThisAddin.Application.ActiveDocument.Unprotect("PASSWORD")
            '
            'Set cropRect = CH_ImgMgr.insertCropRect(objShp, shp)
            cropRect = Me.insertCropRect(sect, objShp, NewPic)

            frm = New frm_pictControl2()
            frm.shp_ImageToBeClipped = NewPic
            frm.cropRect = cropRect
            frm.shp_toBeFilled = objShp
            frm.LayoutName = Me.strFormLayoutName

            frm.Top = 250
            frm.Left = 250
            '
            frm.parentImageMgr = Me
            frm.Show()
            '
            'NewPic.ZOrder(MsoZOrderCmd.msoSendToBack)
            'NewPic.ZOrder(MsoZOrderCmd.msoSendBehindText)
            '
            'Now re-apply dimensions

            Globals.ThisAddin.Application.ScreenUpdating = True
            Globals.ThisAddin.Application.ScreenRefresh()
            '

            GoTo finis

        Catch ex As Exception
            Globals.ThisAddin.Application.ActiveDocument.Undo(1)
            Globals.ThisAddin.Application.ScreenUpdating = True
            Globals.ThisAddin.Application.ScreenRefresh()

        End Try
        '
finis:
        Return NewPic
        '
        '
        '

    End Function

    '
    Public Function imageReplace(InsertOrPaste As String, ByRef sect As Word.Section) As Word.Shape
        Dim hf As Word.HeaderFooter
        Dim shp As Word.Shape
        Dim listOfShapes As Collection
        Dim picDialog As Word.Dialog
        Dim objShp As cShapeMgr
        Dim rng As Word.Range
        Dim NewPic As Word.Shape
        Dim cropRect As Word.Shape
        Dim frm As frm_pictControl2
        '
        Me.currentSect = sect
        imageReplace = Nothing
        NewPic = Nothing
        '
        On Error GoTo finis                                 'Almost invariably due to no data on the clipboard or inappropriate format
        listOfShapes = Me.getPicturePlaceHolders(sect)
        '
        If listOfShapes.Count = 0 Then
            'There were no Cover page images on this page
            MsgBox("Insert/Paste Picture in  only supported in Cover pages with picture placeholders")
            Exit Function
        End If
        '
        hf = sect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        rng = hf.Range
        '
        'picDialog = Dialogs(wdDialogInsertPicture)
        picDialog = Globals.ThisAddin.Application.Dialogs(Word.WdWordDialog.wdDialogInsertPicture)

        objShp = listOfShapes.Item(1)
        Me.strPicturePlaceHolderName = objShp.name
        '
        Select Case InsertOrPaste
            Case "pasteImage"
                'Call hf.rng.PasteSpecial(, , wdFloatOverText, , wdPasteMetafilePicture)
                Call rng.PasteSpecial(, , WdOLEPlacement.wdFloatOverText, , WdPasteDataType.wdPasteMetafilePicture)
                'HFRange.PasteSpecial Placement:=wdFloatOverText, DataType:=wdPasteMetafilePicture
                NewPic = rng.ShapeRange(rng.ShapeRange.Count)
                'Set NewPic = HFSHapes.Item(HFSHapes.Count)
                NewPic.Top = Me.objTools.tools_math_MillimetersToPoints(297.0#) - NewPic.Height - 40.0#
                NewPic.Name = objShp.name
        '
            Case "insertImage"
                picDialog.name = "*.*"
                If picDialog.Display = 0 Then
                    Globals.ThisAddin.Application.ActiveDocument.Undo(1)
                    Exit Function
                End If
                If picDialog.name = "" Then
                    Globals.ThisAddin.Application.ActiveDocument.Undo(1)
                    Exit Function
                End If
                '
                NewPic = hf.Shapes.AddPicture(picDialog.name, False, True, 0, 0, , , rng)
                NewPic.Top = Me.objTools.tools_math_MillimetersToPoints(297.0#) - NewPic.Height - 40.0#
                'NewPic.name = strShapeName
                NewPic.Name = objShp.name
        End Select
        '
        Call NewPic.ScaleWidth(1.0#, MsoTriState.msoTrue)  'sort of resets the record of width & height?
        Call NewPic.ScaleHeight(1.0#, MsoTriState.msoTrue)
        '
        objShp.width_original = NewPic.Width
        objShp.height_original = NewPic.Height

        Call Me.setImagePageProperties(NewPic)          'Page positional properties
        Call Me.setImageScale(NewPic, sect)
        objShp.scaleFactor_W = objShp.width_original / NewPic.Width
        Call Me.setImagePosition(NewPic, sect)
        '
        imageReplace = NewPic
        '
        If Globals.ThisAddin.Application.ActiveDocument.ProtectionType <> WdProtectionType.wdNoProtection Then Call Globals.ThisAddin.Application.ActiveDocument.Unprotect("PASSWORD")
        '
        'Set cropRect = CH_ImgMgr.insertCropRect(objShp, shp)
        cropRect = Me.insertCropRect(sect, objShp, NewPic)
        frm = New frm_pictControl2()
        frm.shp_ImageToBeClipped = NewPic
        frm.cropRect = cropRect
        frm.shp_toBeFilled = objShp
        frm.LayoutName = Me.strFormLayoutName

        frm.Top = 250
        frm.Left = 250
        '
        frm.parentImageMgr = Me
        frm.Show()
        '
        'Now re-apply dimensions

        Globals.ThisAddin.Application.ScreenUpdating = True
        Globals.ThisAddin.Application.ScreenRefresh()
        '

        Exit Function
        '
finis:
        Globals.ThisAddin.Application.ActiveDocument.Undo(1)
        Globals.ThisAddin.Application.ScreenUpdating = True
        Globals.ThisAddin.Application.ScreenRefresh()
        '
        MsgBox("The clipboard is empty...or it doesn't have a picture of the right format in it")
        '
    End Function
    '
    '
    Public Function insertCropRect(ByRef sect As Section, ByRef objShp As cShapeMgr, ByRef shp As Word.Shape) As Word.Shape
        Dim aspectRatioOriginalImage As Single
        Dim clipHeight As Single, clipWidth As Single
        Dim hf As HeaderFooter
        Dim rng As Word.Range
        '
        hf = objShp.hf
        rng = hf.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        aspectRatioOriginalImage = objShp.aspectRatio       'Me.height / Me.width
        clipHeight = shp.Height
        clipWidth = clipHeight / aspectRatioOriginalImage
        If clipWidth <= shp.Width Then

        Else
            clipWidth = shp.Width
            clipHeight = clipWidth * aspectRatioOriginalImage
        End If
        '
        'Set inserCtropRect = ActiveDocument.Shapes.AddShape(msoShapeRectangle, shp.left, shp.top, objShp.width, objShp.height, shp.Anchor)
        insertCropRect = hf.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, shp.Left, shp.Top, clipWidth, clipHeight, rng)
        insertCropRect.LockAspectRatio = MsoTriState.msoTrue
        insertCropRect.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
        insertCropRect.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
        insertCropRect.LockAnchor = True
        insertCropRect.WrapFormat.AllowOverlap = True
        insertCropRect.WrapFormat.Type = WdWrapType.wdWrapNone
        '
        'insertCropRect.left = shp.left
        'insertCropRect.top = shp.top
        insertCropRect.Fill.Transparency = 0.5
        insertCropRect.Fill.BackColor.RGB = RGB(0, 128, 128)
        '
    End Function
    '
    'Simple paste picture method
    Public Sub PasteAsPicture()
        On Error GoTo finis
        Globals.ThisAddin.Application.Selection.PasteSpecial(DataType:=WdPasteDataType.wdPasteMetafilePicture, Placement:=WdOLEPlacement.wdInLine)
        'Globals.ThisAddin.Application.Selection.PasteSpecial(DataType:=WdPasteDataType.wdPasteDeviceIndependentBitmap, Placement:=WdOLEPlacement.wdInLine)
        'Globals.ThisAddin.Application.Selection.PasteSpecial(DataType:=WdPasteDataType.wdPasteEnhancedMetafile, Placement:=WdOLEPlacement.wdInLine)
        'Globals.ThisAddin.Application.Selection.PasteSpecial(DataType:=WdPasteDataType.wdPasteBitmap, Placement:=WdOLEPlacement.wdInLine)
        'Globals.ThisAddin.Application.Selection.PasteSpecial(DataType:=WdPasteDataType.wdPasteHTML, Placement:=WdOLEPlacement.wdInLine)

        Exit Sub
finis:
        MsgBox("There is no picture file on your clipboard")
    End Sub
    '
    Public Sub setImagePageProperties(ByRef shp As Word.Shape)
        Call shp.ScaleWidth(1.0#, MsoTriState.msoTrue)  'sort of resets the record of width & height?
        Call shp.ScaleHeight(1.0#, MsoTriState.msoTrue)
        '
        shp.LockAspectRatio = MsoTriState.msoFalse
        shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
        shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
        shp.LockAnchor = True
        shp.WrapFormat.AllowOverlap = True
        shp.WrapFormat.Type = WdWrapType.wdWrapNone
        '
    End Sub
    '
    Public Sub setImagePosition(ByRef shp As Word.Shape, ByRef sect As Word.Section)
        'This positions the target image in the centre of the page
        Dim AspectRatio_Page As Single
        '
        AspectRatio_Page = sect.PageSetup.PageHeight / sect.PageSetup.PageWidth
        '
        If AspectRatio_Page > 1 Then  'Page is Portrait
            shp.Left = (595 / 2) - shp.Width / 2
            'shp.top = (841 / 2) - shp.height / 2
            shp.Top = (842 - shp.Height) / 2                     ' Sit at the bottom of the page - Allow 20pts for footer
        ElseIf AspectRatio_Page <= 1 Then 'Page is Landscape
            shp.Left = (841 / 2) - shp.Width / 2
            shp.Top = (595 / 2) - shp.Height / 2
        End If
        '
    End Sub
    '
    Public Sub setImageScale(ByRef shp As Word.Shape, ByRef sect As Word.Section)
        ' ASPECT RATIO CONTEXT = HEIGHT/WIDTH
        'set shape size according to page size
        Dim AspectRatio_Page As Single
        Dim AspectRatio_InsPic As Single
        '
        AspectRatio_Page = sect.PageSetup.PageHeight / sect.PageSetup.PageWidth
        AspectRatio_InsPic = shp.Height / shp.Width

        If AspectRatio_Page > 1 Then  'Page is Portrait
            'Aspect Ratio of Inserted shape < 1  = Portrait
            If AspectRatio_InsPic < 1 Then
                'fix width of inserted shape as proportion of portrait page
                shp.Width = 0.8 * sect.PageSetup.PageWidth
                shp.Height = shp.Width * AspectRatio_InsPic
                'Aspect Ratio of Inserted shape >1  = Landscape
            ElseIf AspectRatio_InsPic > 1 Then
                'fix height of inserted shape as proportion of portrait page
                shp.Height = 0.8 * sect.PageSetup.PageHeight
                shp.Width = shp.Height / AspectRatio_InsPic
            End If
        ElseIf AspectRatio_Page <= 1 Then 'Page is Landscape
            '2. Aspect Ratio of Inserted shape > 1
            If AspectRatio_InsPic < 1 Then
                'fix width of inserted shape as proportion of landscape page
                shp.Width = 0.8 * sect.PageSetup.PageWidth
                shp.Height = shp.Width * AspectRatio_InsPic
            ElseIf AspectRatio_InsPic > 1 Then
                'fix height of inserted shape as proportion of portrait page
                shp.Height = 0.8 * sect.PageSetup.PageHeight
                shp.Width = shp.Height / AspectRatio_InsPic
            End If
        End If
    End Sub

End Class
