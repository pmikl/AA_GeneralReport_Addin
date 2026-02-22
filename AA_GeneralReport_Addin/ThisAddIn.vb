Imports Word = Microsoft.Office.Interop.Word
Public Class ThisAddIn
    Private WithEvents myTimer As New System.Windows.Forms.Timer()
    Private WithEvents wordApp As Word.Application
    Private hasHandledFirstDoc As Boolean = False
    Private _lastDoc As Word.Document = Nothing

    Public frm_colorPicker02 As frm_colorPicker02
    '
    Public ribbon As Office.IRibbonUI
    Public strActualDirTemplates As String
    Public point_PriorClick As Drawing.Point                    'To be used in frm_SelectedDocs. We'll show it here. Allows us to remeber before 


    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Dim objGlobals As New cGlobals()
        Dim objTimer As New cTimer()
        Dim objFileMgr As New cFileHandler()
        '
        '
        'Event handler. Other handler are added with document_new and document_open. Note that the
        'SelectionChanged Event Handler is application wide and not docuemtn specific (i.e. it works for all open documents)
        AddHandler Globals.ThisAddIn.Application.DocumentOpen, AddressOf wordApp_DocumentOpen
        AddHandler Globals.ThisAddIn.Application.NewDocument, AddressOf wordApp_DocumentNew
        AddHandler Globals.ThisAddIn.Application.DocumentBeforeClose, AddressOf wordApp_DocumentBeforeClose
        AddHandler Globals.ThisAddIn.Application.DocumentBeforePrint, AddressOf wordApp_DocumentBeforePrint
        AddHandler Globals.ThisAddIn.Application.WindowSelectionChange, AddressOf wrdApp_SelectionChanged
        AddHandler Globals.ThisAddIn.Application.WindowActivate, AddressOf wordApp_WindowActivate
        AddHandler Globals.ThisAddIn.Application.WindowDeactivate, AddressOf wordApp_WindowDeActivate
        '
        '
        wordApp = Globals.ThisAddIn.Application
        Me.frm_colorPicker02 = Nothing

        Me.strActualDirTemplates = ""
        Me.point_PriorClick = Nothing                  'To be used in frm_SelectedDocs. We'll show it here. Allows us to remeber before 
        '
        'Set up the templates directory.. Go for C:\Templates first, then fall back to 'Documents\aa_Templates'
        'The actual directory used for this session is stored in Me.strActualDirTemplates
        '
        '***Verified 20250909
        '
        Dim targetPath As String = "C:\Templates\AA GeneralReport.dotx"
        '
        If System.IO.File.Exists(targetPath) Then
            System.IO.File.Delete(targetPath)
            '
        End If
        '
        Me.strActualDirTemplates = objGlobals.glb_setDir_Templates()
        objFileMgr.file_set_templateFromResources(Me.strActualDirTemplates)
        '
        ' Check if a document is already open (e.g., launched via File Explorer), but before events
        ' were wired up.. Remember asynchronous environment
        '
        If wordApp.Documents.Count > 0 AndAlso Not hasHandledFirstDoc Then
            handleFirstDocumentOpen(wordApp.Documents(1))
        End If
        '
        '
        '****
        'Try  Private Sub rbn_aa_Addin00_Load istead of timer. This event fires when the ribbon is loaded
        '
        'Timer delay to allow the ribbon to load so that my Pgs tab activation code will work
        myTimer.Interval = 2000 ' 1.5 seconds delay
        AddHandler myTimer.Tick, AddressOf on_timer_tick
        myTimer.Start()
        '
    End Sub
    '
    ''' <summary>
    ''' Will handle the situation with lower speed machines where the wiring of the events
    ''' does not finish before the document is opened.. Typically happens when you open
    ''' direct from file explorer
    ''' </summary>
    ''' <param name="doc"></param>
    Private Sub handleFirstDocumentOpen(ByVal doc As Word.Document)
        If Not hasHandledFirstDoc Then
            hasHandledFirstDoc = True
            '
            'MsgBox($"First document opened: {doc.Name}")

            ' Your logic here
            Me.wordApp_DocumentOpen(doc)
            '
        End If
    End Sub
    '
    Private Sub wordApp_WindowActivate(ByVal myDoc As Word.Document, ByVal wndow As Word.Window)
        Dim objStylesMgr As New cStylesManager()
        Dim objCtrls As New cControlsMgr()
        '
        'MsgBox("Window Activated")
        '**** Check for 'AA document status' and adjust tabs
        'objStylesMgr.glb_doc_checkDocType_ActivateTab(objCtrls._strTabId_PagesAndSections)
        '
        If _lastDoc Is Nothing OrElse Not Object.ReferenceEquals(_lastDoc, myDoc) Then
            objStylesMgr.glb_screen_stopRefresh()
            '
            objStylesMgr.glb_doc_checkDocType_ActivateTab(objCtrls._strTabId_PagesAndSections)
            _lastDoc = myDoc
            '
            objStylesMgr.glb_screen_update(True)
        End If


    End Sub
    '
    Private Sub wordApp_WindowDeActivate(ByVal myDoc As Word.Document, ByVal wndow As Word.Window)
        'MsgBox("Window De-Activated")
    End Sub
    '
    Private Sub wordApp_DocumentOpen(ByVal Doc As Word.Document)
        Dim vstoDoc As Microsoft.Office.Tools.Word.Document = Globals.Factory.GetVstoObject(Doc)
        Dim objGlobals As New cGlobals()
        Dim objStylesMgr As New cStylesManager()
        Dim objCtrls As New cControlsMgr()
        Dim objWrdOptions As New cWordOptions()
        Dim objMsgMgr As New cMessageManager()
        Dim objFileMgr As cFileHandler
        Dim strMsg As String = ""
        '
        Dim objProp As New cPropertyMgr()
        '
        objWrdOptions.wrdOptions_set_pasteToInline()
        objWrdOptions.wrdOptions_set_fieldShading("always")
        '
        AddHandler vstoDoc.ActivateEvent, AddressOf Document_Activated
        'AddHandler Globals.ThisAddIn.Application.WindowSelectionChange, AddressOf wrdApp_SelectionChanged

        ' Optional: Activate tab on document open
        Try
            objGlobals.ctrl_tab_Activate(objGlobals._strTabId_PagesAndSections)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Failed to activate tab_aa_PagesAndSections: " & ex.Message)
        End Try
        '
        '
        '**** Check for 'AA document status' and adjust tabs
        objStylesMgr.glb_doc_checkDocType_ActivateTab(objCtrls._strTabId_PagesAndSections)
        '
        'If objGlobals.glb_doc_isAAStdDoc(Doc) Then
        'Doc.AttachedTemplate = objGlobals.glb_getTmpl_FullName()
        'End If
        '
        If objProp.prps_CustomProperty_Exists(Doc, "_AssemblyName") Then
            'If objMsgMgr.msgMgr_dd() Then
            objFileMgr = New cFileHandler()
            '
            objFileMgr.file_get_saveTimeStampedCopy(Doc, "")
            '
            objProp.prps_del_customProperty("_AssemblyLocation", Doc)
            objProp.prps_del_customProperty("_AssemblyName", Doc)
            '
            '
            Doc.AttachedTemplate = objGlobals.glb_getTmpl_FullName()
            '
            Doc.Save()
            '
            '"'filename-yyyyMMdd-hhmmss'"
            MsgBox("The document you just opened has been closed and a copy made which is compatible with the new ACIL Allen Word Addin." + vbCrLf + vbCrLf + "It is this renamed copy, in the format 'filename-yyyyMMdd-hhmmss', which is open on your desktop. It has been saved in the same location as your original document.")
            '
            'MsgBox("Conversion is complete, save/close/re-open and continue working")
            '
            'Doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
            'Else
            'MsgBox("Conversion to AA Addin compatibility has been cancelled")
            'End If

        End If
        '
    End Sub
    '
    Private Sub wordApp_DocumentNew(ByVal Doc As Word.Document)
        Dim vstoDoc As Microsoft.Office.Tools.Word.Document = Globals.Factory.GetVstoObject(Doc)
        Dim objLegal As New cLegalAndAbout()
        Dim objCaptionsMgr As New cCaptionManager()
        Dim objWrdOptions As New cWordOptions()
        Dim objStylesMgr As New cStylesManager()
        Dim objPrint As New cPrintAndDisplayServices()
        Dim objCtrls As New cControlsMgr()
        Dim objThmMgr As New cThemeMgr()
        '
        '
        AddHandler vstoDoc.ActivateEvent, AddressOf Document_Activated

        Try
            'objLegal.insert_CopyRightYear(Doc)
            '
            'Not nessary becuase the Addin carries the most up to date template
            'objStylesMgr.style_upgrade_TemplateStyles()

        Catch ex As Exception
            MsgBox("Failed to pdate the copyright statement. If this problem persists please contact your IT manager")
        End Try
        '
        '
        Try
            objWrdOptions.wrdOptions_set_pasteToInline()
            objWrdOptions.wrdOptions_set_fieldShading("always")
            '
            'Update the captions
            objCaptionsMgr.deleteCustomCaptions()
            objCaptionsMgr.installCustomCaptions()
            '
            objThmMgr.thm_Set_ThemeToAAStd_fromFile(Doc)
            '
        Catch ex As Exception
            MsgBox("Failed to update the custom captions. If this problem persists please contact your IT manager")
        End Try
        '
        objPrint.colour_display_ToEasyView(Globals.ThisAddIn.Application.ActiveDocument)
        '
        objStylesMgr.glb_doc_checkDocType_ActivateTab(objCtrls._strTabId_PagesAndSections)
        '

        '
    End Sub

    '
    Private Sub wrdApp_SelectionChanged(ByVal Sel As Word.Selection)
        Dim objStylesMgr As New cStylesManager()
        Dim objCtrls As New cControlsMgr()
        '
        '
        '**** Check for 'AA document status' and adjust tabs
        If objStylesMgr.glb_doc_isAAStdDoc(Sel.Document) Then
            'MsgBox("Selection Changed in AADoc")
            'objCtrls.ctrl_tabSet_Visibility("all")
            'Globals.Ribbons.rbn_Styles.mnu_grpReport_Columns.Visible = True
            'Globals.Ribbons.rbn_Styles.RibbonUI.ActivateTab(strTabId)
        Else
            'MsgBox("Selection Changed in Not AADoc")
            '
            'objCtrls.ctrl_tabSet_Visibility("all", False)
            'Globals.Ribbons.rbn_Styles.RibbonUI.ActivateTabMso("TabHome")
            'Globals.Ribbons.rbn_Styles.RibbonUI.ActivateTabMso("TabHome")
            'Globals.Ribbons.rbn_Styles.RibbonUI.ActivateTabMso("TabHome")
            '
        End If
        '
        '
    End Sub
    '
    Private Sub Document_Activated(sender As Object, e As Microsoft.Office.Tools.Word.WindowEventArgs)
        Dim objStylesMgr As New cStylesManager()
        Dim objCtrls As New cControlsMgr()
        '
        '**** Check for 'AA document status' and adjust tabs
        'MsgBox("Document Activated")
        'objStylesMgr.glb_doc_checkDocType_ActivateTab(objCtrls._strTabId_PagesAndSections)
        '
    End Sub
    '
    Private Sub wordApp_DocumentBeforePrint(ByVal Doc As Word.Document, ByRef cancel As Boolean)
        Dim objPrint As New cPrintAndDisplayServices()
        '
        'objPrint.print_with_ColourChange(Doc)
        objPrint.colour_display_ToDesignView(Doc)
        '
        cancel = False
        '
    End Sub
    '
    ''' <summary>
    ''' We de-activate the controls if thare are no open documents
    ''' </summary>
    ''' <param name="Doc"></param>
    ''' <param name="cancel"></param>
    Private Sub wordApp_DocumentBeforeClose(ByVal Doc As Word.Document, ByRef cancel As Boolean)
        'Dim objCtrls As New cControlsMgr()
        '
        'If Globals.ThisAddIn.Application.Documents.Count <= 1 And Not cancel Then
        'objCtrls.ctrl_tabSet_Visibility("all", False)
        'End If
        '
    End Sub


    Private Sub HookActivateEvent(doc As Word.Document)
        Try
            Dim vstoDoc As Microsoft.Office.Tools.Word.Document = Globals.Factory.GetVstoObject(doc)
            AddHandler vstoDoc.ActivateEvent, AddressOf Document_Activated
        Catch ex As Exception
            ' Optional: log or handle cases where VSTO object can't be retrieved
        End Try
    End Sub
    '
    Public Sub on_timer_tick(s As Object, args As EventArgs)
        Dim timer As New System.Windows.Forms.Timer()
        Dim objGlobals As New cGlobals()
        '
        myTimer.Stop()
        '
        Try
            'objGlobals.ctrl_tab_Activate(objGlobals._strTabId_PagesAndSections)
        Catch ex As Exception
            'System.Diagnostics.Debug.WriteLine("Failed to activate tab_aa_PagesAndSections: " & ex.Message)
        End Try
        '
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
