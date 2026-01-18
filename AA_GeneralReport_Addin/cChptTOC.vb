Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cChptTOC
    Inherits cChptBase
    Public Sub New()
        MyBase.New()
    End Sub
    '
    Public Function is_TOCPage(ByRef sect As Word.Section)
        Return False
    End Function


End Class
