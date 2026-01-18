Imports System.Windows.Forms
Imports System.IO
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core


Public Class cCloneMgr
    Public mySrcDoc As Word.Document
    Public myDestDoc As Word.Document
    '
    Public Sub New(ByRef srcDoc As Word.Document, ByRef destDoc As Word.Document)
        Me.mySrcDoc = srcDoc
        Me.myDestDoc = destDoc
    End Sub
    '
    Public Sub New()

    End Sub
    '
    Public Sub clone_Doc_byCopy(ByRef srcDoc As Word.Document, ByRef destDoc As Word.Document)
        'Dim sect As Word.Section
        '
        srcDoc.Content.Copy()
        destDoc.Content.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting)
        '
        '
        srcDoc.Saved = True
        srcDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
        '
    End Sub
    '
    Public Function cloneDoc() As Boolean
        Dim rslt As Boolean
        Dim destDoc As Word.Document
        '
        rslt = True
        destDoc = Globals.ThisAddIn.Application.ActiveDocument
        Me.myDestDoc = destDoc
        '
        Try
            If Not IsNothing(Me.mySrcDoc) Then
                'Me.cloneDoc(Me.mySrcDoc, Me.myDestDoc, doContents)
                Me.clone_Doc_byCopy(Me.mySrcDoc, Me.myDestDoc)
                rslt = True
            Else
                rslt = False
            End If

        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
    End Function

    '
    Public Sub cloneDoc(ByRef srcDoc As Word.Document, ByRef destDoc As Word.Document, doContents As Boolean)
        Dim srcSect, destSect As Word.Section
        Dim tmpl As Word.Template
        Dim objSectMgr As New cSectionMgr()
        Dim doHeaderFooters As Boolean
        Dim i As Integer
        Dim rng, rngSrc As Word.Range
        '
        doHeaderFooters = True
        '
        tmpl = srcDoc.AttachedTemplate
        '
        Me.copyStyles(tmpl, destDoc)
        '
        'objSectMgr.deleteAllSections(destDoc)
        '
        For i = 1 To srcDoc.Sections.Count
            srcSect = srcDoc.Sections.Item(i)
            If i = 1 Then
                destSect = destDoc.Sections.Item(i)
            Else
                destSect = destDoc.Sections.Add(, Word.WdSectionStart.wdSectionNewPage)
            End If
            Me.cloneSection(srcSect, destSect, doHeaderFooters)
            rng = destSect.Range
            rng.Collapse(Word.WdCollapseDirection.wdCollapseStart)
            rng.Paragraphs.Add()
            rng.Paragraphs.Add()
            '
            rng = destSect.Range
            rng.Collapse(Word.WdCollapseDirection.wdCollapseStart)

        Next

        For i = 1 To srcDoc.Sections.Count
            srcSect = srcDoc.Sections.Item(i)
            rngSrc = srcSect.Range
            rngSrc.MoveEnd(WdUnits.wdParagraph, -1)
            '
            '*** range copy is just not working
            rngSrc.Copy()
            'rngSrc.Copy()
            '*** 
            '
            'rngSrc.Select()
            'rngSrc.Copy()
            'Globals.ThisAddIn.Application.Selection.Copy()

            If i = 1 Then
                'destSect = destDoc.Sections.Item(i)
            Else
                'destSect = destDoc.Sections.Add(, Word.WdSectionStart.wdSectionNewPage)
            End If
            'Me.cloneSection(srcSect, destSect, doHeaderFooters)

            destSect = destDoc.Sections.Item(i)

            '
            rng = destSect.Range
            rng.Collapse(Word.WdCollapseDirection.wdCollapseStart)
            rng.Paragraphs.Add()
            rng.Paragraphs.Add()
            '

            'Me.cloneSection(srcSect, destSect, doHeaderFooters)
            rng = destSect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            If doContents Then
                'rng.Paste()
                'rng.PasteAndFormat(WdRecoveryType.wdUseDestinationStylesRecovery)
                rng.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting)

            End If

        Next
        '
    End Sub
    '
    Public Sub cloneSection(ByRef srcSection As Section, ByRef destSection As Section, doHeaderFooters As Boolean)
        'This method will reproduce the settings of the source Section (srcSection) in the
        'destination Section (destSection
        '
        'Dim strOrientation As String
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim hf As Word.HeaderFooter
        '
        'strOrientation = "portrait"
        'If srcSection.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then strOrientation = "landscape"
        '
        destSection.PageSetup.PaperSize = srcSection.PageSetup.PaperSize
        destSection.PageSetup.Orientation = srcSection.PageSetup.Orientation
        'If strOrientation = "portrait" Then
        'destSection.PageSetup.Orientation = WdOrientation.wdOrientPortrait
        'Else
        'destSection.PageSetup.Orientation = WdOrientation.wdOrientLandscape
        'End If
        destSection.PageSetup.GutterPos = srcSection.PageSetup.GutterPos
        destSection.PageSetup.DifferentFirstPageHeaderFooter = srcSection.PageSetup.DifferentFirstPageHeaderFooter
        destSection.PageSetup.OddAndEvenPagesHeaderFooter = srcSection.PageSetup.OddAndEvenPagesHeaderFooter
        destSection.PageSetup.MirrorMargins = srcSection.PageSetup.MirrorMargins
        '
        'Copy page dimensions
        destSection.PageSetup.TopMargin = srcSection.PageSetup.TopMargin
        destSection.PageSetup.LeftMargin = srcSection.PageSetup.LeftMargin
        destSection.PageSetup.BottomMargin = srcSection.PageSetup.BottomMargin
        destSection.PageSetup.RightMargin = srcSection.PageSetup.RightMargin
        destSection.PageSetup.Gutter = srcSection.PageSetup.Gutter
        destSection.PageSetup.HeaderDistance = srcSection.PageSetup.HeaderDistance
        destSection.PageSetup.FooterDistance = srcSection.PageSetup.FooterDistance
        '
        Globals.ThisAddIn.Application.ScreenRefresh()
        '
        If doHeaderFooters Then
            objHfMgr.hf_hfs_linkUnlinkAll(destSection, False)
            objHfMgr.hf_headers_delete(destSection)

            For Each hf In srcSection.Headers
                objHfMgr.hf_hfs_CopyHeader(hf, srcSection, destSection)
            Next
            '
            objHfMgr.hf_footers_delete(destSection)

            For Each hf In srcSection.Footers
                objHfMgr.hf_hfs_CopyFooter(hf, srcSection, destSection)
            Next
        End If

        'Call objHfMgr.hf_hfs_CopyHeaderFooter("header", srcSection, destSection)
        'Call objHfMgr.hf_hfs_CopyHeaderFooter("footer", srcSection, destSection)
        '

    End Sub

    '
#Region "Setup Source and Dest"
    '
    Public Sub setSrcDoc(ByRef srcDoc As Word.Document)
        Me.mySrcDoc = srcDoc
    End Sub
    '
    Public Sub setDestDoc(ByRef destDoc As Word.Document)
        Me.myDestDoc = destDoc
    End Sub
    '
    Public Function isOKToClone() As Boolean
        Dim rslt As Boolean
        '
        rslt = True
        '
        Try
            If Me.myDestDoc.Name = Me.mySrcDoc.Name Then rslt = False
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
    End Function
    '

#End Region
    '
#Region "Styles"
    '
    Public Sub copyStyles(src As Word.Template, dest As Word.Document)
        '
        dest.CopyStylesFromTemplate(src.FullName)
        '
        dest.Saved = True
        '
    End Sub
    '

    Public Sub copyStyles(src As Word.Document, dest As Word.Document)
        Dim normalStyle As Word.Style
        '
        normalStyle = src.Styles.Item("Normal")
        dest.CopyStylesFromTemplate(src.AttachedTemplate)
        '
        dest.Save()
        dest.Saved = True
        '
    End Sub
    '
    Public Function styleExists(styleName As String, ByRef myDoc As Word.Document) As Boolean
        Dim rslt As Boolean
        Dim myStyle As Word.Style
        Dim x As Single
        '
        rslt = False
        '
        Try
            myStyle = myDoc.Styles.Item(styleName)
            x = myStyle.Font.Size
            rslt = True
        Catch ex As Exception
            rslt = False
        End Try

        Return rslt
    End Function
#End Region

End Class
