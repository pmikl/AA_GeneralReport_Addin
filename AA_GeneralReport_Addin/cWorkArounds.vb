Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cWorkArounds
    Inherits cGlobals
    Public Sub New()
        MyBase.New()
    End Sub
    '
    ''' <summary>
    ''' 20250710.. Problems when inserting sections. They are there, but the references to the
    ''' new sections are not correct (they point to older sections).. It's as though the internal
    ''' software hasn't yet caught up with the fact that there are new sections.. The WordAround
    ''' is to force the document to check it's list of sections.
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub wrk_fix_forSectionProblem(ByRef sect As Word.Section, Optional waitInMilliseconds As Integer = 250)
        Dim numSections As Integer
        Dim objTimer As New cTimer()
        '
        objTimer.tmr_waitHere_milliseconds(waitInMilliseconds) '
        '
        numSections = sect.Range.Document.Sections.Count
        numSections = sect.Range.Document.Sections.Count

    End Sub
    '
    Public Sub wrk_fix_forSectionProblem(ByRef myDoc As Word.Document, Optional waitInMilliseconds As Integer = 250)
        Dim numSections As Integer
        Dim objTimer As New cTimer()
        '
        objTimer.tmr_waitHere_milliseconds(waitInMilliseconds) '
        '
        numSections = myDoc.Sections.Count
        numSections = myDoc.Sections.Count

    End Sub

    '
    ''' <summary>
    ''' The cursor will continuoulsy flick between I beam, default and other settings.. After some
    ''' operations. Operations that trigger this seem to chnage over time.. It appears that fields in the
    ''' Headers and or footers can trigger this continuous polling action. It seems that by removing the
    ''' offending fields and then replacing the race/polling condition disappears
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub wrk_fix_forCursorRace_Alt(ByRef sect As Word.Section)
        Dim objHFMgr As New cHeaderFooterMgr()
        '*** To get rid of race condition... Seems to work 20231219
        'Remove footer to get rid of fields, then put it back
        '
        objHFMgr.hf_footers_delete(sect)
        objHFMgr.hf_footers_insert(sect)
        '
    End Sub
    '
    Public Sub wrk_fix_forCursorRace()
        '
        Me.glb_flds_updateStyleRefsFooters()
        '
    End Sub

End Class
