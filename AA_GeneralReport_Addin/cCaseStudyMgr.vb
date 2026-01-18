Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word
Public Class cCaseStudyMgr
    Inherits cPlHBase

    Public Sub New()
        MyBase.New()

    End Sub
    '
    '
    ''' <summary>
    ''' This method will determine if the current section is a full 'page' Case Study section. It does so by looking for the
    ''' style tag_caseStudy in the primary Header.
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function cst_is_caseStudySection(ByRef sect As Section) As Boolean
        Dim rng As Range
        Dim rngStyle As Word.Style
        Dim rslt As Boolean
        '
        rslt = False
        rng = sect.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary).Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rngStyle = rng.Style
        '
        'MsgBox("Style is = " + rngStyle.NameLocal)
        '
        'If drCell.Range.Style Is Globals.ThisDocument.Application.ActiveDocument.Styles("tag_coverPage") Then
        If rngStyle.NameLocal = "tag_caseStudy" Then
            rslt = True
        End If
        '
        Return rslt
    End Function
    '
    Public Function cst_insert_fullPageCaseStudy(ByRef rng As Word.Range) As String
        Dim objSectMgr As New cSectionMgr()
        Dim objMsgMgr As New cMessageManager()
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objBnrMgr As New cChptBanner()
        Dim sect As Word.Section
        Dim strMsg, strTagStyle As String
        '
        strMsg = ""
        sect = rng.Sections.Item(1)
        '
        'Get the tag style of the parent section. If we can't find a tag style, then we defualt to chapter body and hope for the best
        'We wont use it for the moment becase case study sections will take on the tag styles of their parent
        '
        strTagStyle = objHFMgr.hf_tags_getTagStyleName(sect, "primaryOrFirstPage")
        If strTagStyle = "" Then strTagStyle = objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_chpt_body)
        '
        If Not objSectMgr.sct_Sel_IsIn_Or_JustUnderTable() Then
            'rng = objGlobals.glb_get_wrdSelRng()
            If rng.Sections.Item(1).PageSetup.Orientation = WdOrientation.wdOrientPortrait Then
                sect = objSectMgr.sct_insert_SectionBounded(rng, "cstudy_Prt", 6, "newPage", False)
            Else
                sect = objSectMgr.sct_insert_SectionBounded(rng, "cstudy_Lnd", 6, "newPage", False)
            End If
            '
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'myStyle = rng.Document.Styles.Item("Heading (CaseStudy)")
            'rng.Paragraphs.Item(1).Style = rng.Document.Styles.Item("Heading (CaseStudy)")
            'rng.Text = "CASE STUDY"
            'rng.Select()
            'rng = Me.Plh_Captions_InsertCaptions("CaseStudy", rng, True)
            'rng.Font.Size = myStyle.Font.Size
            '
            rng = Me.cst_insert_Caption(rng)
            rng.Select()

        Else
            strMsg = "inTable"
        End If
        '
        Return strMsg
        '
    End Function
    '
    ''' <summary>
    ''' This method will insert the Case Study numbered heading caption
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <returns></returns>
    Public Function cst_insert_Caption(ByRef rng As Word.Range) As Word.Range
        Dim myStyle As Word.Style
        '
        myStyle = rng.Document.Styles.Item("Heading (CaseStudy)")
        '
        rng = Me.Plh_Captions_InsertCaptions("CaseStudy", rng, True)
        rng.Font.Size = myStyle.Font.Size
        '
        Return rng
    End Function
    '
    ''' <summary>
    ''' This method inserts a Case study table (edge to edge) to allow for partial page case studies
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <returns></returns>
    Public Function cst_insert_partialPageCaseStudy(ByRef rng As Word.Range) As Word.Table
        Dim tbl As Word.Table
        Dim drCell As Word.Cell
        Dim para As Word.Paragraph
        Dim rngP As Word.Range
        Dim myStyle As Word.Style
        '
        myStyle = rng.Document.Styles.Item("Heading (CaseStudy)")
        '
        tbl = MyBase.Plh_insert_PlaceHolder_WithTest(rng, "CaseStudy_HalfPage")
        '
        drCell = tbl.Range.Cells.Item(1)
        para = drCell.Range.Paragraphs.Item(1)
        rngP = para.Range
        rngP.Font.Size = myStyle.Font.Size

        rng = tbl.Range.Cells.Item(2).Range
        rng.Text = "Insert case study text here"
        rng.MoveEnd(WdUnits.wdCharacter, -1)
        '
        Return tbl
        '
    End Function


End Class
