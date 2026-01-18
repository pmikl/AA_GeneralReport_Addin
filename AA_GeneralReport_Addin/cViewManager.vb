Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word

Public Class cViewManager
    Inherits cGlobals
    Public Sub New()
        MyBase.New()
    End Sub
    '
    ''' <summary>
    ''' This method will change the view of the Active Document to 1 column / 1 row if the
    ''' orientation of the current section (sect) is Landscape, or to 2 columns / 1 row if it is Portrait
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub vw_change_ColumnsAndRows(ByRef sect As Word.Section, Optional zoomPercent As Integer = 75)
        Dim myDoc As Word.Document
        '
        'The page view was touchy when I set columns and rows, so leaving it as fit full page
        '
        myDoc = glb_get_wrdActiveDoc()
        'myDoc.ActiveWindow.View.Zoom.PageFit = WdPageFit.wdPageFitFullPage
        '
        'GoTo finis
        ' 
        ' 
        If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then
            myDoc.ActiveWindow.View.Zoom.PageColumns = 2
            myDoc.ActiveWindow.View.Zoom.PageRows = 1
        Else
            myDoc.ActiveWindow.View.Zoom.PageColumns = 3
            myDoc.ActiveWindow.View.Zoom.PageRows = 1

        End If
        myDoc.ActiveWindow.View.Zoom.Percentage = zoomPercent
        '
finis:
    End Sub
    '
    ''' <summary>
    ''' This method will change the view of the Active Document to 1 column / 1 row if the
    ''' orientation of the current section (sect) is Landscape, or to 2 columns / 1 row if it is Portrait
    ''' </summary>
    Public Sub vw_change_toPageFitBestFit(ByRef myDoc As Word.Document, Optional zoomPercent As Integer = 69)
        '
        myDoc.ActiveWindow.View.Zoom.PageFit = WdPageFit.wdPageFitBestFit
        myDoc.ActiveWindow.View.Zoom.Percentage = zoomPercent
        ' 
    End Sub
    '
    '
    ''' <summary>
    ''' This method will change the view of the Active Document to 1 column / 1 row if the
    ''' orientation of the current section (sect) is Landscape, or to 2 columns / 1 row if it is Portrait
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub vw_change_toMultiplePages(ByRef sect As Word.Section, Optional zoomPercent As Integer = 75)
        Dim myDoc As Word.Document
        '
        myDoc = glb_get_wrdActiveDoc()
        myDoc.ActiveWindow.View.Zoom.PageFit = WdPageFit.wdPageFitNone
        myDoc.ActiveWindow.View.Zoom.Percentage = zoomPercent
        'myDoc.ActiveWindow.View.Zoom.PageFit = 
        ' 
    End Sub
    '

    '
    '

    ''' <summary>
    ''' This method will set the view to the specified number of columns and rows on
    ''' a page
    ''' </summary>
    ''' <param name="columns"></param>
    ''' <param name="rows"></param>
    Public Sub vw_change_ColumnsAndRows(columns As Integer, rows As Integer, Optional zoomPercent As Integer = 75)
        Dim myDoc As Word.Document
        '
        myDoc = glb_get_wrdActiveDoc()
        glb_view_setToPrintLayout()
        ' 
        myDoc.ActiveWindow.View.Zoom.PageColumns = columns
        myDoc.ActiveWindow.View.Zoom.PageRows = rows
        '
        myDoc.ActiveWindow.View.Zoom.Percentage = zoomPercent
        '
    End Sub
    '
    Public Sub vw_fit_fullPage(ByRef sect As Word.Section, Optional zoomPercent As Integer = 75)
        Dim myDoc As Word.Document
        '
        myDoc = glb_get_wrdActiveDoc()
        myDoc.ActiveWindow.View.Zoom.PageFit = WdPageFit.wdPageFitFullPage
        '
        myDoc.ActiveWindow.View.Zoom.Percentage = zoomPercent
        '
    End Sub
    '


End Class
