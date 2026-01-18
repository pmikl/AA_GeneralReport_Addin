Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
''' <summary>
''' 
'''Originally written in vba, some account taken for conversion, but this
'''was not a priority at the time this class was written
'''
'''Peter Mikelaitis October 2015...http://mikl.com.au
'''Ported to VB.NET Jan 2017 from version 97p21p05
'''
''' </summary>
Public Class cFormatMgr
    Public name As String
    Public myDoc As Document
    Public sect As Section
    Public objSectMgr As cSectionMgr

    Public Sub New()
        Me.name = "hello"
        'Me.myDoc = Globals.ThisDocument.Application.ActiveDocument
        Me.objSectMgr = New cSectionMgr()
        Me.myDoc = objSectMgr.objGlobals.glb_get_wrdActiveDoc()
    End Sub
    '
    Public Sub frmt_ClearTabs(ByRef myStyle As Style)
        Dim i As Integer
        Dim tb As TabStop
        '
        For i = myStyle.ParagraphFormat.TabStops.Count To 1 Step -1
            tb = myStyle.ParagraphFormat.TabStops(i)
            tb.Clear()
        Next i
        '
    End Sub

End Class
