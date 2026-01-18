Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cTableStyles
    Inherits cStylesManager
    Public Sub New()
        MyBase.New()
    End Sub
    '


    ''' <summary>
    ''' This method will insert the table style 'aac Table (no lines)' if it doesn't exist. If
    ''' it does it applies the appropriate formatting, which gives us an 'in' to changing this
    ''' construct at a later date. The doExtraFormatting option is applied to an existing style
    ''' if doExtraFormatting is true. It is always applied (regardless of the option setting)
    ''' to the style if it is newly created
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="doExtraFormatting"></param>
    Public Function tblstyl_add_aacTableNoLines(ByRef myDoc As Word.Document, Optional doExtraFormatting As Boolean = False) As Word.Style
        Dim styl As Word.Style
        '
        styl = glb_tbl_getAACTableNoLinesStyle(myDoc)
        '
        Return styl
    End Function
    '
    ''' <summary>
    ''' The act of getting the style forces it's creation if it doesn't exist
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="doExtraFormatting"></param>
    ''' <returns></returns>
    Public Function tblstyl_add_aacTableBasic(ByRef myDoc As Word.Document, Optional doExtraFormatting As Boolean = False) As Word.Style
        Dim styl As Word.Style
        '
        styl = glb_tbl_getAACTableBasicStyle(myDoc)
        '
        Return styl
    End Function


End Class
