Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core

''' <summary>
'''Ported to VB.NET 17th Jan 2017 from version 97p21p05
''' </summary>
''' 
Public Class cMessageManager
    Public Sub New()

    End Sub
    '
    Public Sub msgMgr_dlg_legacyWaterMarks()
        Dim strMsg As String
        '
        strMsg = "Your document may have no water marks or" + vbCrLf _
            + "just legacy water marks." + vbCrLf + vbCrLf _
            + "To fix this, just apply or re-apply a new set of water marks."
        '
        MsgBox(strMsg,, "Missing or legacy water marks")
        '
        '
    End Sub
    '
    Public Function msgMgr_dd() As Boolean
        Dim dlgResult As Integer
        Dim strMsg, strMsg2 As String
        Dim rslt As Boolean

        strMsg = "This is an 'old format' document and, subject to your approval" + vbCrLf _
                 + "will be made compatible with the 'Ribbon Addin'. The steps are:" + vbCrLf + vbCrLf _
                 + "1." + vbTab + "The file you opened will be resaved at the same location," + vbCrLf _
                 + vbTab + "(preserving the original) and will be renamed to ," + vbCrLf _
                 + vbTab + "'filename-yyyyMMdd-hhmmss'" + vbCrLf + vbCrLf _
                 + "2." + vbTab + "The legacy ribbon elements will be stripped out of this" + vbCrLf _
                 + vbTab + "resaved copy." + vbCrLf + vbCrLf _
                 + "3." + vbTab + "The converted resaved copy will be remain open on" + vbCrLf _
                 + vbTab + "your desktop." + vbCrLf + vbCrLf _
                 + "4." + vbTab + "Rename it as required, resave and continue working with" + vbCrLf _
                 + vbTab + "the modified copy. It will not trigger this conversion process" + vbCrLf _
                 + vbTab + "again and will be responsive to the new Addin Ribbon." + vbCrLf + vbCrLf _
                 + "5." + vbTab + "It is good practise to close/re-open your new document," + vbCrLf _
                 + vbTab + "just to ensure that these changes stick, and that any" + vbCrLf _
                 + vbTab + "lingering artefacts are removed." + vbCrLf + vbCrLf _
                + "Do you wish to continue?"

        strMsg2 = "This is an 'old format' document and, subject to your approval" + vbCrLf _
                 + "will be made compatible with the 'Ribbon Addin'. The steps are:" + vbCrLf + vbCrLf _
                 + "1." + vbTab + "The file you just opened will have the legacy ribbon" + vbCrLf _
                 + vbTab + "elements stripped from it." + vbCrLf + vbCrLf _
                 + "2." + vbTab + "After this process is finished, it is good practise" + vbCrLf _
                 + vbTab + "to close/re-open the document, just to ensure that these" + vbCrLf _
                 + vbTab + "changes stick, and that any lingering artefacts stay removed." + vbCrLf + vbCrLf _
                 + "3." + vbTab + "If you are concerned, cancel this process and make a" + vbCrLf _
                 + vbTab + "backup of your document and then try again." + vbCrLf + vbCrLf _
                 + "Do you wish to continue?"

        rslt = False
        dlgResult = MsgBox(strMsg, vbYesNo + vbDefaultButton2, "Change over to AA Addin")
        If dlgResult = vbYes Then rslt = True
        '
        Return rslt

    End Function

    Public Function msgMgr_dlg_doDocToWCAGExported() As Boolean
        Dim dlgResult As Integer
        Dim strMsg As String
        Dim rslt As Boolean
        '
        strMsg = "The current document is about to be converted to an 'Accessible'." + vbCrLf _
            + "format as determined by Word's inbuilt 'Accessibility'checker" + vbCrLf _
            + "(see the Review tab)." + vbCrLf + vbCrLf _
            + "Your original document will be automatically saved and then renamed" + vbCrLf _
            + "in the current directory." + vbCrLf + vbCrLf _
            + "It is this renamed document that will be converted. If the orignal name" + vbCrLf _
            + "was 'filename.docx', then the name of the converted file will take" + vbCrLf _
            + "the form 'filename-wcag-yyyymmdd-hhmmss" + vbCrLf + vbCrLf _
            + "Conversion typically takes between 30 to 40 seconds for a medium sized document." + vbCrLf + vbCrLf _
            + "Do you wish to continue..?"
        '
        rslt = False
        dlgResult = MsgBox(strMsg, vbYesNo + vbDefaultButton2, "Active Document to WCAG")
        If dlgResult = vbYes Then rslt = True
        '
        Return rslt
    End Function
    '
    '
    Public Function msgMgr_dlg_themesStylesAndTemplateWarning() As Boolean
        Dim dlgResult As Integer
        Dim strMsg As String
        Dim rslt As Boolean
        '
        strMsg = "This menu item will update the styles and theme to match those" + vbCrLf _
            + "of a standard ACIL Allen document. It will also attach the standard " + vbCrLf _
            + "ACIL Allen template, allowing you to use the ribbon functions." + vbCrLf + vbCrLf _
            + "Doing this to a large imported document does not make this an" + vbCrLf _
            + "ACIL Allen document as it will be missing meta data used by the" + vbCrLf _
            + "ribbon functions." + vbCrLf + vbCrLf _
            + "Consequently the functions may not behave as expected, but some will." + vbCrLf _
            + "Tread carefully when using this function (backup your document)." + vbCrLf + vbCrLf _
            + "Do you wish to continue..?"
        '
        rslt = False
        dlgResult = MsgBox(strMsg, vbYesNo + vbDefaultButton2, "Attach template and update styles and theme")
        If dlgResult = vbYes Then rslt = True
        '
        Return rslt
    End Function
    '    
    '
    Public Function msgMgr_dlg_cloneDocumentWarning() As Boolean
        Dim dlgResult As Integer
        Dim strMsg As String
        Dim rslt As Boolean
        '
        strMsg = "This menu item will copy the entire contents of the current (Active)" + vbCrLf _
            + "document (source document) to a blank document (destination" + vbCrLf _
            + "document) opened from the standard ACIL Allen template." + vbCrLf + vbCrLf _
            + "When finished the source document will be closed (without any" + vbCrLf _
            + "changes), leaving the new 'cloned' document open and ready for editing." + vbCrLf + vbCrLf _
            + "Doing this to a large imported document does not make it an" + vbCrLf _
            + "ACIL Allen document, as it will be missing meta data used by the " + vbCrLf _
            + "ribbon functions. But the Addin will now recognise the document" + vbCrLf _
            + "as probably ACIL Allen and make the ribbon functions available." + vbCrLf + vbCrLf _
            + "Some ribbon functions may NOT BEHAVE AS EXPECTED, but generally" + vbCrLf _
            + "wont be destructive, just non-responsive." + vbCrLf + vbCrLf _
            + "Tread carefully when editing your new document with the ribbon" + vbCrLf _
            + "(i.e. back it up)." + vbCrLf + vbCrLf _
            + "Do you wish to continue..?"
        '
        rslt = False
        dlgResult = MsgBox(strMsg, vbYesNo + vbDefaultButton2, "Attach template And update styles And theme")
        If dlgResult = vbYes Then rslt = True
        '
        Return rslt
    End Function
    '    

    '
    Public Function msgMgr_dlg_fillAllTableHeaders() As Boolean
        Dim dlgResult As Integer
        Dim strMsg As String
        Dim rslt As Boolean
        '
        strMsg = "This menu item will fill the header rows of all 'Regular/Standard'," + vbCrLf _
            + "tables in the document with the colour of your choice. " + vbCrLf + vbCrLf _
            + "It is a good idea to save your document before doing this, as you" + vbCrLf _
            + "can only undo it by using this same function to apply another colour" + vbCrLf + vbCrLf _
            + "Do you wish to continue..?"
        '
        rslt = False
        dlgResult = MsgBox(strMsg, vbYesNo + vbDefaultButton2, "Fill all Table Header Rows")
        If dlgResult = vbYes Then rslt = True
        '
        Return rslt
    End Function
    '    '

    '
    Public Function msgMgr_dlg_stylesToBlack() As Boolean
        Dim dlgResult As Integer
        Dim strMsg As String
        Dim rslt As Boolean
        '
        strMsg = "This menu item will set all styles in the document to black." + vbCrLf + vbCrLf _
            + "If you need to undo this and you have already saved the changes, " + vbCrLf _
            + "go to the 'Styles Tab' and select 'Update fron template' in the 'Reset styles' panel" + vbCrLf + vbCrLf _
            + "Do you wish to continue..?"
        '
        rslt = False
        dlgResult = MsgBox(strMsg, vbYesNo + vbDefaultButton2, "Active Document to WCAG")
        If dlgResult = vbYes Then rslt = True
        '
        Return rslt
    End Function
    '    '

    '
    Public Function msgMgr_dlg_doDocToWCAG() As Boolean
        Dim dlgResult As Integer
        Dim strMsg As String
        Dim rslt As Boolean
        '
        strMsg = "The current document is about to be converted to WCAG compliance." + vbCrLf + vbCrLf _
            + "Please note that this is a one way trip and may take a minute or so" + vbCrLf + vbCrLf _
            + "Do you wish to continue..?"
        '
        rslt = False
        dlgResult = MsgBox(strMsg, vbYesNo + vbDefaultButton2, "Active Document to WCAG")
        If dlgResult = vbYes Then rslt = True
        '
        Return rslt
    End Function
    '    '
    Public Sub UpdateFunctionIsFinished()
        Dim dlgResult As Integer

        dlgResult = MsgBox("The update function has finished",, "AAC Update")

    End Sub
    '
    '
    Public Sub pageNumbers_Fault_ToESBodyChangeFailed(strType As String)
        Dim strMsg As String
        '
        strMsg = "This is probably because the section is set to the two digit (i.e. X-Y)" + vbCrLf _
            + "Appendix numbering format... You'll need to use the" + vbCrLf _
            + "'Page Numbering dialogue' to change the format back to the single digit format" + vbCrLf + vbCrLf _
            + "Select the 'Display the page numbering dialog' at the bottom of the" + vbCrLf _
            + "'Page # Formatting' menu...Then uncheck 'include chapter number'"
        '
        Select Case strType
            Case "ES"
                strMsg = "The attempted change to 'Executive Summary' numbering has failed." + vbCrLf + vbCrLf + strMsg
            Case "body"
                strMsg = "The attempted change to 'Report body' numbering has failed." + vbCrLf + vbCrLf + strMsg
            Case Else
        End Select

        MsgBox(strMsg,, "Page number reformatting fault")

    End Sub

    '
    Public Sub pageNumbers_Fault_ChangeFailed(strTypeOfAction As String)
        Dim strMsg As String
        '
        strMsg = "The attempt to change the page numbering format has failed." + vbCrLf + vbCrLf _
            + "If you need to change the format of an 'X-Y' numbering scheme" + vbCrLf _
            + "(i.e. as found in the Appendices) to single digits you'll need" + vbCrLf _
            + "to use the 'Page Numbering dialogue' to uncheck 'include chapter' number" + vbCrLf + vbCrLf _
            + "See the 'Pages and Sections>Finalise Report (far right)>Page # formatting"
        '
        Select Case strTypeOfAction
            Case "reformat"
            Case "resetToPortrait"
                strMsg = "Reset to Portrait and Reset to Landscape both attempt to take the" + vbCrLf _
                    + "page numbering format back to the single digit Report body standard" + vbCrLf + vbCrLf + strMsg
            Case Else
        End Select

        MsgBox(strMsg,, "Page number reformatting fault")

    End Sub
    '

    '
    Public Function doTableColumnUndoMessage() As Boolean
        Dim dlgResult As Integer
        Dim strMsg As String
        '

        strMsg = "When the AA Table column/row functions were used, or when the" + vbCrLf _
            + "'Copy Table' function was used, a copy of the original Table was placed." + vbCrLf _
            + "on the ClipBoard." + vbCrLf + vbCrLf _
            + "If you choose to proceed, the Table holding your cursor will be replaced" + vbCrLf _
            + "with the contents of the ClipBoard." + vbCrLf + vbCrLf _
            + "Do you wish to continue..?"


        '
        doTableColumnUndoMessage = False
        dlgResult = MsgBox(strMsg, vbYesNo + vbDefaultButton2, "Finalise")
        If dlgResult = vbYes Then doTableColumnUndoMessage = True
        '
    End Function
    '
    Public Function msg_stdTbl_hasNoCaption() As Boolean
        Dim dlgResult As Integer
        Dim strMsg As String
        Dim rslt As Boolean
        '
        strMsg = "This table has no caption." + vbCrLf + vbCrLf _
            + "Do you wish to continue..?"
        '
        rslt = False
        dlgResult = MsgBox(strMsg, vbYesNo + vbDefaultButton2, "No Caption")
        If dlgResult = vbYes Then rslt = True
        '
        Return rslt
    End Function
    '
    '
    Public Function msg_stdTbl_hasNoSourceNoteRow() As Boolean
        Dim dlgResult As Integer
        Dim strMsg As String
        Dim rslt As Boolean
        '
        strMsg = "This table has no Source/Note paragraph." + vbCrLf + vbCrLf _
            + "Do you wish to continue..?"
        '
        rslt = False
        dlgResult = MsgBox(strMsg, vbYesNo + vbDefaultButton2, "No Source/Note")
        If dlgResult = vbYes Then rslt = True
        '
        Return rslt
    End Function
    '
    '
    Public Function msg_tbl_doPasteTable() As Boolean
        Dim dlgResult As Integer
        Dim strMsg As String
        Dim rslt As Boolean
        '
        strMsg = "Your cursor was not in a table." + vbCrLf + vbCrLf + "So the contents of the ClipBoard will be pasted at the current position." + vbCrLf + vbCrLf _
            + "Do you wish to continue..?"
        '
        rslt = False
        dlgResult = MsgBox(strMsg, vbYesNo + vbDefaultButton2, "No Source/Note")
        If dlgResult = vbYes Then rslt = True
        '
        Return rslt
    End Function
    '

    '
    '
    Public Function doFinaliseMessage() As Boolean
        Dim dlgResult As Integer
        Dim strMsg As String
        '
        strMsg = "The document Is about to be finalised. You'll be notified when its finished." + vbCrLf + vbCrLf _
            + "Please note that this can take a while if you have a large, complex document" + vbCrLf + vbCrLf _
            + "Do you wish to continue..?"
        '
        doFinaliseMessage = False
        dlgResult = MsgBox(strMsg, vbYesNo + vbDefaultButton2, "Finalise")
        If dlgResult = vbYes Then doFinaliseMessage = True
    End Function
    '
    Public Function docCopyInstruction() As System.Windows.Forms.DialogResult
        docCopyInstruction = MsgBox("The automated process consists of the following steps;" & Constants.vbCr & Constants.vbCr _
            & "1.   You'll be asked to select one or more documents to convert/import." & Constants.vbCr _
            & "      A useful result is only obtained if these document(s) are based on" & Constants.vbCr _
            & "      the original vba template 'Acil Allen Template.dotm'" & Constants.vbCr & Constants.vbCr _
            & "2.   The 'Import' software will create a new empty 'Destination" & Constants.vbCr _
            & "      document' based on the .NET template (It will be invisible" & Constants.vbCr _
            & "      so as not to clutter the screen)" & Constants.vbCr & Constants.vbCr _
            & "3.   If you click 'Import', the software will do a section by section copy" & Constants.vbCr _
            & "      of the contents of the 'Source Document' into the 'Destination" & Constants.vbCr _
            & "      document' of item 2." & Constants.vbCr & Constants.vbCr _
            & "4.   The Destination document will then be saved to an " & Constants.vbCr _
            & "      'AAC conversions' sub folder in your 'Documents' folder" & Constants.vbCr & Constants.vbCr _
            & "5.   Documents may take between 5-20 seconds to convert" & Constants.vbCr & Constants.vbCr _
            & "6.   The software will notify you when the import process is complete" & Constants.vbCr & Constants.vbCr _
            & "7.   A red document name indicates the import encountered a problem" & Constants.vbCr & Constants.vbCr _
            & "When asked, left click on the bottom button to close the Dialog", ".Net Export Information", MsgBoxStyle.OkOnly)
    End Function

    Public Sub tableSpacerRowConversionError()
        Dim strMsg As String
        '
        strMsg = "Tables could not be adjusted to use 'spacer_tbl' style" & vbCrLf
        strMsg = strMsg & "It is likely that you didn't pick up the style in" & vbCrLf
        strMsg = strMsg & "a previous template update. Please see your Admin Staff. Or.." & vbCrLf & vbCrLf
        strMsg = strMsg & "Select File->Options->Addins->Templates (from the Manage Dropdown list)"
        strMsg = strMsg & "->Click Go Button->Check 'Automatically update document styles'" & vbCrLf & vbCrLf
        strMsg = strMsg & "Then close and re-open your document and try again"
        '
        MsgBox(strMsg)
        '
    End Sub
    '
    '
    Public Sub notYetImplemented()
        Dim strMsg As String
        Dim dlgResult As Integer
        '
        strMsg = "This option has not yet been implemented." & vbCr & vbCr _
            & "Due to structural differences between the old and new templates," & vbCr _
            & "it is unlikely that earlier (i.e. current template) mechanisms can be applied directly" & vbCr & vbCr _
            & "So we'll need to review/clarify the underlying functional requirement"
        dlgResult = MsgBox(strMsg)
    End Sub
    '
    Public Sub colourRowsErrorMessage()
        Dim strMsg As String
        Dim dlgResult As Integer
        '
        strMsg = "Make certain that you have selected the rows that you wish to fill." & Constants.vbCr & Constants.vbCr _
            & "This function does not work on irregular tables." & Constants.vbCr _
            & "An irregular table is one where you have joined or split cells so that" & Constants.vbCr & Constants.vbCr _
            & "the standard orthogonal structure no longer exits.. Choose 'Colour Cells' instead"
        dlgResult = MsgBox(strMsg)
        '
    End Sub
    '
    Public Sub IsInBoxFigureRec()
        Dim dlgResult As Integer
        '
        dlgResult = MsgBox("Your cursor is in a AAC Box, Figure, Recommendation" & vbCr _
            & "or Key Finding" & vbCr & vbCr _
            & "This function should only be used to add columns to" & vbCr _
            & "standard AAC Tables." & vbCr & vbCr _
            & "So please move your cursor to an appropriate Table" & vbCr _
            & "and try again", vbOKOnly + vbInformation, "Template Message")

    End Sub
    '
    Public Sub UpdateFigureCaptionsErrorMsg()
        Dim dlgResult As Integer
        '
        dlgResult = MsgBox("The Update/Refresh cannot start" & vbCr & vbCr _
            & "The probable cause is the location of your cursor" & vbCr & vbCr _
            & "Please move your cursor to an empty paragraph" & vbCr _
            & "in the body of the document that is at least one" & vbCr _
            & "paragraph clear of any Tables, Table Structures, " & vbCr _
            & "Boxes, Figures or Banner Headings and try again" & vbCr & vbCr _
            & "Don't place it in the Cover, Contacts Or TOC pages", vbOKOnly + vbInformation, "Template Message")

    End Sub
    '

    Public Sub msg_insertionPoint_IsIn_Or_JustUnderATable()
        Dim dlgResult As Integer
        '
        dlgResult = MsgBox("Your current cursor position Is in  (or next to) a Table." & vbCr & vbCr _
            & "Please move your cursor to a clear" & vbCr _
            & "location in the document and try again", vbOKOnly + vbInformation, "Template Message")
    End Sub
    '
    '

    Public Sub msg_insertionPoint_isInsideATable()
        Dim dlgResult As Integer
        '
        dlgResult = MsgBox("Your current cursor position is inside a Table." & vbCr & vbCr _
            & "Please move your cursor to a clear" & vbCr _
            & "location in the document and try again", vbOKOnly + vbInformation, "Template Message")
    End Sub
    '

    '
    Public Sub cannotInsertSection(strSectionType As String, strTag As String)
        Dim dlgResult As Integer
        Dim strMsg As String
        '
        strMsg = "Cannot insert a " & strSectionType & " in a " & strTag & " Section"
        '
        dlgResult = MsgBox(strMsg, vbOKOnly + vbInformation, "Template Message")
    End Sub
    '
    Public Sub cannotAdjustHeadingWidth()
        Dim dlgResult As Integer
        Dim strMsg As String
        '
        strMsg = "The 'Heading Width' function can only be used on cover pages" & vbCr _
            & "Please make sure that your cursor is in a Cover Page" & vbCr & vbCr _
            & "If you want to adjust the page width, try the" & vbCr _
            & "'Toggle Section Width' Function in the 'Other Section Options' group"
        '
        dlgResult = MsgBox(strMsg, vbOKOnly + vbInformation, "Template Message")
    End Sub
    '
    Public Sub errorInAdjustHeadingWidth()
        Dim dlgResult As Integer
        '
        dlgResult = MsgBox("The Heading area failed to adjust - The most likely cause is" & vbCr _
            & "that a change has been made to the cover page structure." & vbCr _
            & "The quickest solution is to reinsert a new Cover Page." & vbCr & vbCr _
            & "If the problem persists, please contact your local IT support staff for assistance", vbOKOnly + vbExclamation, "Heading Warning")

    End Sub
    '
    Public Sub msg_cropping_Error()
        Dim strMsg As String
        '
        strMsg = "Error in the cropping manager... " + vbCrLf + vbCrLf _
            + "Did you have an image on your clipboard?" + vbCrLf + vbCrLf _
            + "If the error persists you may need to use" + vbCrLf _
            + "the 'Raw Image from file... option" + vbCrLf + vbCrLf _
            + "Just remember to crop your image to same shape as the object" + vbCrLf _
            + "you are trying to fill, otherwise the result will be distorted"
        MsgBox(strMsg)

    End Sub

    Public Sub messageTest()
        MsgBox("The image insert/function has failed. Probable causes are;" & vbCr & vbCr &
            "- Your cover page does not support custom pictures." & vbCr &
            "- Your cover page is suitable, but the target 'image placeholder' " & vbCr &
            "  has been accidently deleted." & vbCr & vbCr &
            "In either case, the quickest solution is to refresh the cover page" & vbCr &
            "by using the 'Select Cover Page Option' under the 'ACIL ALLEN Report' tab on the toolbar", vbInformation, "Template Message")
        Exit Sub
    End Sub
    '
    Public Function deleteReportMessage() As Boolean
        Dim dlgResult As Integer
        '
        deleteReportMessage = False
        dlgResult = MsgBox("This function will delete ALL sections after the letter" & vbCr _
            & "This action cannot be undone." & vbCr & vbCr _
            & "Do you wish to continue.?", vbYesNo + vbDefaultButton2, "Delete Warning")
        If dlgResult = vbYes Then deleteReportMessage = True

    End Function
    '
    Public Function msg_warning_MetaDataRemoval() As Boolean
        Dim dlgResult As Integer
        Dim rslt As Boolean
        '
        rslt = False
        dlgResult = MsgBox("This function will delete ALL Meta Data from the" + vbCrLf _
            + "current document and then re-saves it in the current folder" + vbCrLf _
            + "under the new name;" + vbCrLf + vbCrLf _
            + "'filename-noMetaData-yyyymmdd-hhmmss.docx'." + vbCrLf + vbCrLf _
            + "Document 'Collaboration' functions depend on Meta data, so" + vbCrLf _
            + "removing it will impact 'shared' experiences, such as Co-Authoring." + vbCrLf + vbCrLf _
            + "This action cannot be undone." & vbCrLf & vbCrLf _
            + "Do you wish to continue.?", vbYesNo + vbDefaultButton2, "Meta Data Removal Warning")
        If dlgResult = vbYes Then rslt = True
        '
        Return rslt
    End Function
    '
    Public Function deleteAllMessage() As Boolean
        Dim dlgResult As Integer
        '
        deleteAllMessage = False
        dlgResult = MsgBox("This function will delete ALL sections and cannot be undone." + vbCr & vbCr _
            + "A message will advise you when the build is complete" + vbCrLf _
            + "Normally this will take between 17 to 50 seconds depending on the power of your machine" + vbCrLf + vbCrLf _
            + "Do you wish to continue.?", vbYesNo + vbDefaultButton2, "Delete Warning")
        If dlgResult = vbYes Then deleteAllMessage = True
        '
    End Function

    Private Function deleteTestMessage() As Boolean
        Dim dlgResult As Integer
        '
        deleteTestMessage = False
        dlgResult = MsgBox("Chapter failed to insert - Most likely cause is a" & vbCr _
            & "missing AutoText Entry for the Chapter. The template has been" & vbCr _
            & "modified or corrupted. Please contact your" & vbCr _
            & "local IT support staff for assistance", vbOKOnly + vbExclamation, "Insert Warning")
        If dlgResult = vbYes Then deleteTestMessage = True

    End Function
    '
    '
    Public Function msgMgr_dlg_doDocToWCAGAllPlaceHolders_toInLine() As Boolean
        Dim dlgResult As Integer
        Dim strMsg As String
        Dim rslt As Boolean
        '
        strMsg = "This functions examines all placeholders in the current" + vbCrLf _
            + "document and ensures that they are all in line" + vbCrLf + vbCrLf _
            + "Do you wish to continue..?"
        '
        rslt = False
        dlgResult = MsgBox(strMsg, vbYesNo + vbDefaultButton2, "Active Document to WCAG")
        If dlgResult = vbYes Then rslt = True
        '
        Return rslt
    End Function
    '
    Public Function msgMgr_msg_tooNearATable() As String
        Dim strMsg As String
        '
        strMsg = "Your cursor must be located at least one clear paragraph away from any existing tables or placeholders, otherwise they'll merge in unexpected ways." + vbCrLf + vbCrLf + "Please relocate your insertion point and try again"
        '
        Return strMsg
    End Function
    '
    Public Function msgMgr_msg_notAvailableInBrief() As String
        Dim strMsg As String
        '
        strMsg = "This function is not available in the ACIL Allen Brief document type."
        '
        Return strMsg

    End Function
End Class
