Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports System.Text
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.Word
Imports System.Drawing
'
'
'rev 01.00  20250830
'
Public Class cFileHandler
    Public objGlobals As New cGlobals()
    Public strFolderPath As String
    Public timeToLive As TimeSpan
    Public _strResavePath As String
    Public _strScratchDir As String = "aa_scratch_wrd"                         'To be used in creation of the timestamped scratch file

    Public Sub New(strFolderPath As String)
        'Office Recovery Files will live for 48 hrs before being deleted
        'Me.timeToLive = New TimeSpan(0, 48, 0, 0, 0)
        '
        Me.strFolderPath = strFolderPath
    End Sub
    '
    Public Sub New(strSaveFolder As String, lifeSpan As TimeSpan)
        'Me._strResavePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\officerecovery\"
        Me._strResavePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + strSaveFolder
        'Me.timeToLive = New TimeSpan(0, 48, 0, 0, 0)
        Me.timeToLive = lifeSpan
    End Sub
    '
    Public Sub New()
        'Office Recovery Files will live for 48 hrs before being deleted
        '
        Me.strFolderPath = Environment.CurrentDirectory
    End Sub
    '
    Public Function file_get_siteInformation() As Collection
        Dim lstOfSiteInfo As Collection
        Dim strUpdateSite As String
        '
        lstOfSiteInfo = objGlobals.glb_get_VersionInformation()
        Try
            strUpdateSite = CStr(lstOfSiteInfo("updateSite"))
        Catch ex As Exception
            strUpdateSite = ""
        End Try
        '
        Return lstOfSiteInfo
    End Function
    '
    Public Function file_delete_File(strFileFullName As String) As Boolean
        Dim rslt As Boolean
        '
        rslt = False
        '
        Try
            If System.IO.File.Exists(strFileFullName) Then
                System.IO.File.Delete(strFileFullName)
                rslt = True
            End If
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
        '
    End Function
    '
    Public Function file_make_dir(strDirectoryFullName As String) As Boolean
        Dim directoryInfo As System.IO.DirectoryInfo
        Dim rslt As Boolean
        '
        'Create the directory, 
        '
        rslt = False
        directoryInfo = New DirectoryInfo(strDirectoryFullName)
        '
        If directoryInfo.Exists Then
            rslt = True
        Else
            directoryInfo = System.IO.Directory.CreateDirectory(strDirectoryFullName)
            If directoryInfo.Exists Then
                rslt = True
            End If
        End If
        '
        Return rslt
        '
    End Function
    '
    Public Function file_get_RptExampleFromResources(strRptType As String) As String
        Dim objGlobals As New cGlobals()
        Dim objScratchMgr As New cFileScratchMgr()
        Dim docSourceFile As Byte()
        Dim strFileName, strFileFullName, strActualDirTemplates As String
        '
        strFileName = ""
        strFileFullName = ""
        docSourceFile = Nothing
        '
        Select Case strRptType
            Case "Prt"
                strFileName = "AA_ReportExample_Prt.dotx"
                docSourceFile = My.Resources.AA_ReportPrt_Example
                '
            Case "Lnd"
                strFileName = "AA_ReportExample_Lnd.dotx"
                docSourceFile = My.Resources.AA_ReportLnd_Example
            Case "Brf"
                strFileName = "AA_ReportExample_Brf.dotx"
                docSourceFile = My.Resources.AA_ReportBrf_Example
                '
                '
                '**** temporary to build new guides from old documents.. Making certina there is
                'either no references or the appropriate reference
            Case "stylesGuide"
                strFileName = "StylesGuide-exampleDoc.docx"
                'docSourceFile = My.Resources.AA_StylesGuide
                docSourceFile = My.Resources.AA_ReportPrt_Example

            Case "stylesGuide_AccessibleAware"
                strFileName = "StylesGuide-exampleDoc-AccessibleAware.docx"
                'docSourceFile = My.Resources.AA_StylesGuide_AccessibleAware
                docSourceFile = My.Resources.AA_ReportPrt_Example

        End Select

        strActualDirTemplates = objScratchMgr.scratch_get_scratchDirectory()
        If Me.file_make_dir(strActualDirTemplates) Then
            'strActualDirTemplates = strActualDirTemplates + "\"
            strFileFullName = strActualDirTemplates + "\" + strFileName
            File.WriteAllBytes(strFileFullName, docSourceFile)
        End If


        Return strFileFullName
        '
    End Function
    '
    ''' <summary>
    ''' This method will delete the specified directory and its contents
    ''' </summary>
    ''' <param name="strDirectoryFullName"></param>
    Public Sub file_delete_Directory(strDirectoryFullName As String)
        Dim directoryInfo As System.IO.DirectoryInfo
        Dim recursive As Boolean = True
        '
        Try
            directoryInfo = New DirectoryInfo(strDirectoryFullName)
            '
            If directoryInfo.Exists Then
                'Delete direcory and contents
                System.IO.Directory.Delete(strDirectoryFullName, True)
            End If

        Catch ex As Exception

        End Try
        '
    End Sub

    ''' <summary>
    ''' This method will download the General Report template (AA GeneralReport) from Resources
    ''' into the specified directory. Normally this is the actual directory avaialble from
    ''' objGlobals.glb_getDir_TemplatesActual... Every time the Addin is opened, the existence
    ''' of the directory is checked and if it is missing it is re-created. Note that strActualDirTemplates
    ''' is in the format C:\Templates... Internally it uses C:\Templates\, the extra '\' is added inside
    ''' this method
    ''' </summary>
    ''' <param name="strActualDirTemplates"></param>
    ''' <returns></returns>
    Public Function file_set_templateFromResources(strActualDirTemplates As String, Optional doTimeStamp As Boolean = False) As Boolean
        Dim objGlobals As New cGlobals()
        Dim docSourceFile As Byte()
        Dim strFileName, strFileFullName As String
        Dim rslt As Boolean
        '
        strFileName = ""
        strFileFullName = ""
        docSourceFile = Nothing
        '
        strFileName = objGlobals.glb_var_TemplateFileName
        docSourceFile = My.Resources.AA_GeneralReport
        Try
            strFileFullName = objGlobals.glb_getDir_inUseforTemplates + "\" + strFileName
            File.WriteAllBytes(strFileFullName, docSourceFile)
            rslt = True
            '
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
        '        '
    End Function
    '
    '
    Public Function file_get_templateFromWeb(Optional strFileId As String = "", Optional strWebLocation As String = "acilallen.com.au") As String
        Dim strFileFullPath As String
        '
        strFileFullPath = Me.file_get_resourcesFromWeb(Me.objGlobals.glb_var_TemplateFileName, strFileId, strWebLocation)
        '
        Return strFileFullPath
        '
    End Function
    '
    Public Function file_get_resourcesFromResource(strResourcFileType As String, strnewFileName As String, strFileId As String) As String
        Dim docSourceFile As Byte()
        Dim dlgFolder As FolderBrowserDialog
        Dim dlgRslt As DialogResult
        Dim strFileFullPath As String = ""
        Dim tokens() As String

        '
        Select Case strResourcFileType
            Case "AA_GeneralReport.dotx"
                docSourceFile = My.Resources.AA_GeneralReport
                tokens = strnewFileName.Split(".")
                strnewFileName = tokens(0) + ".dotx"
                '
            Case "AA_ReportPrt"
                docSourceFile = My.Resources.AA_ReportPrt_Example
                tokens = strnewFileName.Split(".")
                strnewFileName = tokens(0) + ".docx"
                '
            Case "AA_ReportLnd"
                docSourceFile = My.Resources.AA_ReportLnd_Example
                tokens = strnewFileName.Split(".")
                strnewFileName = tokens(0) + ".docx"
                '
            Case "AA_ReportBrf"
                docSourceFile = My.Resources.AA_ReportBrf_Example
                tokens = strnewFileName.Split(".")
                strnewFileName = tokens(0) + ".docx"
                '
            Case "AA_ThemeForRpt"
                docSourceFile = My.Resources.AA_Theme_for_GeneralReport_with_CustClrs_20240808
                tokens = strnewFileName.Split(".")
                strnewFileName = tokens(0) + ".thmx"
                '
            Case "AA_StylesGuide"
                'docSourceFile = My.Resources.AA_StylesGuide
                'tokens = strnewFileName.Split(".")
                'strnewFileName = tokens(0) + ".docx"
                '
            Case "AA_StylesGuide_Accessible"
                'docSourceFile = My.Resources.AA_StylesGuide_Accessible
                'tokens = strnewFileName.Split(".")
                'strnewFileName = tokens(0) + ".docx"
                '

            Case Else
                docSourceFile = Nothing
        End Select
        '
        dlgFolder = New FolderBrowserDialog()
        dlgFolder.SelectedPath = Globals.ThisAddIn.strActualDirTemplates
        dlgRslt = dlgFolder.ShowDialog()
        '
        If dlgRslt = DialogResult.OK Then
            strFolderPath = dlgFolder.SelectedPath + "\"
            strFileFullPath = Me.file_get_newFileName(strnewFileName, strFolderPath, strFileId)
            '
            Try
                File.WriteAllBytes(strFileFullPath, docSourceFile)
            Catch ex As Exception
                strFileFullPath = ""
            End Try
            '
        Else
            strFileFullPath = "cancel"
        End If
        '
        Return strFileFullPath
        '
    End Function

    '
    ''' <summary>
    ''' Downloads resources from the web. The full web path used is 'strFullWebPath = strWebPath + strFileName' where strWebPath is selected
    ''' according to the value of strWebLocation.. If it's 'acilallen.com.au', then we go to the acil allen site. If it's 'mikl.net.au' we go to
    ''' 'mikl.net.au'
    ''' </summary>
    ''' <param name="strFileName"></param>
    ''' <param name="strFileId"></param>
    ''' <param name="strWebLocation"></param>
    ''' <returns></returns>
    Public Function file_get_resourcesFromWeb(strFileName As String, strFileId As String, Optional strWebLocation As String = "acilallen.com.au") As String
        Dim objWCAG As New cWCAGMgr()
        Dim objBBMgr As New cBBlocksHandler()
        Dim dlgFolder As FolderBrowserDialog
        Dim dlgRslt As DialogResult
        Dim client As System.Net.WebClient
        Dim strSoftwareType As String
        Dim strWebPath As String
        Dim strFullWebPath As String
        Dim strFileFullPath As String
        Dim strFolderPath As String
        Dim myFileInfo As FileInfo
        'strFileName = "AA GeneralReport.dotx"
        '
        'Me.file_get_newFileName()

        strSoftwareType = objGlobals.glb_get_SoftwareType()
        '
        'Use C:\Templates as first choice. If not available, then place it in the Documents folder under \AA Resources
        'strFolderPath = "C:\Templates\"
        strFolderPath = objGlobals.glb_getDir_inUseforTemplates + "\"
        '
        If Not Directory.Exists(strFolderPath) Then
            strFolderPath = objGlobals.glb_getDir_documentsLocal() + "\AA Resources\"
            If Not Directory.Exists(strFolderPath) Then
                Directory.CreateDirectory(strFolderPath)
            End If
        End If
        '
        strFileFullPath = Me.file_get_newFileName(strFileName, strFolderPath, strFileId)
        'strFileFullPath = strFolderPath + strFileName
        myFileInfo = New FileInfo(strFileFullPath)
        '
        dlgFolder = New FolderBrowserDialog()
        dlgFolder.SelectedPath = myFileInfo.DirectoryName()
        dlgRslt = dlgFolder.ShowDialog()
        '
        If dlgRslt = DialogResult.OK Then
            strFileFullPath = dlgFolder.SelectedPath + "\" + myFileInfo.Name
            '
            strWebPath = ""
            Select Case strWebLocation
                Case "acilallen.com.au"
                    Select Case strSoftwareType
                        Case "addin"
                            If strFileName Like "*.docx" Then
                                strWebPath = "http://templates.acilallen.com.au/word/GeneralReport_Addin/resources/"
                            Else
                                strWebPath = ""
                            End If
                        Case "template"
                            If strFileName Like "*.docx" Then
                                strWebPath = "http://templates.acilallen.com.au/word/GeneralReport/resources/"
                            Else
                                strWebPath = "http://templates.acilallen.com.au/word/GeneralReport/install/"
                            End If
                    End Select
                    System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.SystemDefault
                    strFullWebPath = strWebPath + strFileName
                    '
                    client = New System.Net.WebClient()
                    client.DownloadFile(strFullWebPath, strFileFullPath)
                    '
                Case "mikl.net.au"
                    Select Case strSoftwareType
                        Case "addin"
                            If strFileName Like "*.docx" Then
                                strWebPath = "https://mikl.net.au/org_aa/office/word/GeneralReport_Addin/resources/"
                            Else
                                strWebPath = ""
                            End If
                        Case "template"
                            If strFileName Like "*.docx" Then
                                strWebPath = "https://mikl.net.au/org_aa/office/word/GeneralReport/resources/"
                            Else
                                strWebPath = "https://mikl.net.au/org_aa/office/word/GeneralReport/install/"
                            End If
                    End Select
                    System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
                    strFullWebPath = strWebPath + strFileName
                    '
                    client = New System.Net.WebClient()
                    client.DownloadFile(strFullWebPath, strFileFullPath)
                    '
                Case Else
                    Select Case strSoftwareType
                        Case "addin"
                            If strFileName Like "*.docx" Then
                                strWebPath = "http://templates.acilallen.com.au/word/GeneralReport_Addin/resources/"
                            Else
                                strWebPath = ""
                            End If
                        Case "template"
                            If strFileName Like "*.docx" Then
                                strWebPath = "http://templates.acilallen.com.au/word/GeneralReport/resources/"
                            Else
                                strWebPath = "http://templates.acilallen.com.au/word/GeneralReport/install/"
                            End If
                    End Select
                    System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.SystemDefault
                    strFullWebPath = strWebPath + strFileName
                    '
                    client = New System.Net.WebClient()
                    client.DownloadFile(strFullWebPath, strFileFullPath)
                    '
            End Select
            '
        ElseIf dlgRslt = DialogResult.Cancel Then
            strFileFullPath = "cancel"
            '
        End If
        '
        '
        Return strFileFullPath
        '
    End Function


    Public Sub file_get_imageFromWeb(ByRef rng As Word.Range)
        Dim objWCAG As New cWCAGMgr()
        Dim objBBMgr As New cBBlocksHandler()
        Dim objBBlk As BuildingBlock
        Dim objTemplate As Word.Template
        Dim iShp As InlineShape
        Dim strWebPath As String
        Dim strFullWebPath As String
        Dim strImageName As String
        Dim strFileFullPath As String
        Dim strFolderPath As String
        Dim myFileInfo As FileInfo
        '
        strWebPath = "http://templates.acilallen.com.au/word/images/"
        strImageName = "aac_pict_indigenous_00.png"
        strImageName = "\" + strImageName

        strFullWebPath = strWebPath + strImageName
        '
        'strFolderPath = "C:\Templates\"
        strFolderPath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\AAC Images"
        '
        If Not Directory.Exists(strFolderPath) Then
            Directory.CreateDirectory(strFolderPath)
        End If
        '
        strFileFullPath = strFolderPath + strImageName
        myFileInfo = New FileInfo(strFileFullPath)
        '
        Try
            If Not myFileInfo.Exists Then
                My.Computer.Network.DownloadFile(strFullWebPath, strFileFullPath)
            End If
            '
            'img = New Bitmap(strFileFullPath)
            'Clipboard.SetImage(img)
            '
            iShp = rng.InlineShapes.AddPicture(strFileFullPath)
            objWCAG.wcag_alttext_write("Indigenous artwork", iShp)
            'objWCAG.wcag_set_decorative(iShp, True)
            '
            objBBlk = objBBMgr.getBuildingBlockEntry("aac_pict_indigenous_00", "ContactsPage", WdBuildingBlockTypes.wdTypeCustom1)
            '
            If IsNothing(objBBlk) Then
                objTemplate = rng.Document.AttachedTemplate
                objBBlk = objTemplate.BuildingBlockEntries.Add("aac_pict_indigenous_00", WdBuildingBlockTypes.wdTypeCustom1, "ContactsPage", iShp.Range)
                objTemplate.Save()
                objTemplate.Saved = True
            End If
            'rng.Paste()
            'j = rng.InlineShapes.Count
            'k = rng.ShapeRange.Count
            '
        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' This method will insert the local file 'strLocalFileFullPath' to the InLine shapes of the range rng. The method
    ''' will return the inline shape. Note that if iShp is NOT nothing then it can be converted to Floating using
    ''' iShp.ConvertToShape
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="width"></param>
    ''' <param name="strLocalFileFullPath"></param>
    ''' <returns></returns>
    Public Function file_insert_imageFromFile(ByRef rng As Word.Range, width As Single, strLocalFileFullPath As String) As InlineShape
        Dim iShp As InlineShape
        iShp = Nothing
        '
        Try
            If Not (strLocalFileFullPath = "") Then
                iShp = rng.InlineShapes.AddPicture(strLocalFileFullPath)
                iShp.LockAspectRatio = True
                iShp.Width = width
                '
            End If
        Catch ex As Exception
            iShp = Nothing
        End Try
        '
        Return iShp
        '
    End Function
    '
    ''' <summary>
    ''' This method is meant to be an alternative to file_insert_imageFromFile. In this case we insert any image (img)
    ''' supplied and return it as an inLine SHape. Typically img is obtained form My.Resources. So we have no reliance
    ''' on external file structures
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="width"></param>
    ''' <param name="img"></param>
    ''' <returns></returns>
    Public Function file_insert_imageFromResources(ByRef rng As Word.Range, width As Single, img As Image) As InlineShape
        Dim objImgMgr As New cImageMgr()
        Dim iShp As InlineShape
        Dim j As Integer
        iShp = Nothing
        '
        j = 0
        Clipboard.SetImage(img)
        '
        iShp = Nothing
        Try
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            '*** The inLine option doesn't seem to have any affect 20231124. SO we have to convert to an inline shape ourselves
            rng.PasteSpecial(DataType:=WdPasteDataType.wdPasteBitmap, Placement:=WdOLEPlacement.wdInLine)
            'rng.MoveEnd(WdUnits.wdParagraph, 2)
            '
            'iShp = objImgMgr.img_get_ImageAsInlineShape(rng)
            'iShp.LockAspectRatio = True
            'iShp.Width = width
            '
            'rng.MoveEnd(WdUnits.wdParagraph, -1)
            '
        Catch ex As Exception
            iShp = Nothing
        End Try
        '
        Return iShp


    End Function
    '
    '
    ''' <summary>
    ''' This method is meant to be an alternative to file_insert_imageFromFile. In this case we insert any image (img)
    ''' supplied and return it as an inLine SHape. Typically img is obtained form My.Resources. So we have no reliance
    ''' on external file structures
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="drCell"></param>
    ''' <param name="img"></param>
    ''' <returns></returns>
    Public Function file_insert_imageFromResources(ByRef rng As Word.Range, ByRef drCell As Word.Cell, img As Image) As InlineShape
        Dim objImgMgr As New cImageMgr()
        Dim iShp As InlineShape
        Dim width As Single
        Dim j As Integer
        iShp = Nothing
        '
        j = 0
        Clipboard.SetImage(img)
        width = drCell.Width - drCell.LeftPadding - drCell.RightPadding
        '
        iShp = Nothing
        Try
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            '*** The inLine option doesn't seem to have any affect 20231124. SO we have to convert to an inline shape ourselves
            rng.PasteSpecial(DataType:=WdPasteDataType.wdPasteBitmap, Placement:=WdOLEPlacement.wdInLine)
            'rng.MoveEnd(WdUnits.wdParagraph, 2)
            '
            rng = drCell.Range
            iShp = objImgMgr.img_get_ImageAsInlineShape(drCell)
            iShp.LockAspectRatio = True
            iShp.Width = width
            '
            'rng.MoveEnd(WdUnits.wdParagraph, -1)
            '
        Catch ex As Exception
            iShp = Nothing
        End Try
        '
        Return iShp


    End Function
    '


    Public Sub file_get_templateFromWeb_xx()
        Dim strWebPath As String
        Dim strFolderPath As String
        Dim strFileName As String
        Dim strFileFullPath As String
        Dim myFileInfo As FileInfo
        '
        'https://docs.microsoft.com/en-us/dotnet/visual-basic/developing-apps/programming/computer-resources/how-to-download-a-file
        '
        strFolderPath = "C:\Templates\"
        'strFileName = "GeneralReport_test.dotx"
        strFileName = "GeneralReport-test.dotx"
        strFileFullPath = strFolderPath + strFileName
        strWebPath = "http://templates.acilallen.com.au/word/GeneralReport/install/GeneralReport.dotx"
        '
        Try
            My.Computer.Network.DownloadFile(strWebPath, strFileFullPath)
            'My.Computer.Network.DownloadFile(strWebPath, strFileFullPath, "username", "password")
            '
            myFileInfo = New FileInfo("C:\Templates\GeneralReport.dotx")
            '
            Try
                If myFileInfo.Exists Then
                    FileSystem.Rename("C:\Templates\GeneralReport.dotx", "C:\Templates\GeneralReport_old.dotx")
                End If
            Catch ex1 As Exception

            End Try
            '
            '
            myFileInfo = New FileInfo("C:\Templates\GeneralReport-test.dotx")
            '
            Try
                If myFileInfo.Exists Then
                    'FileSystem.Rename("C:\Templates\GeneralReport-test.dotx", "C:\Templates\GeneralReport-test2.dotx")
                    FileSystem.Rename("C:\Templates\GeneralReport-test.dotx", "C:\Templates\GeneralReport.dotx")
                End If
            Catch ex As Exception

            End Try
            '
            '
        Catch ex As Exception

        End Try
        '
        'If Not Directory.Exists(strFolderPath) Then
        'My.Computer.FileSystem.CreateDirectory(strFolderPath)
        'End If

        '
    End Sub
    '
    Public Sub file_doc_toWCAG(ByRef myDoc As Word.Document, Optional doAsNewFile As Boolean = True)
        Dim objMsgMgr As New cMessageManager()
        Dim objWCAGMgr As New cWCAGMgr()
        'Dim myDocInfo As New FileInfo(myDoc.FullName)
        Dim objGlobals As New cGlobals()
        Dim wrdApp As Word.Application
        'Dim objGlobals As New cGlobals()
        Dim strNewFullSavePath As String
        'Dim newDoc, newDoc2 As Word.Document
        '
        wrdApp = objGlobals.glb_get_wrdApp
        '
        '
        If doAsNewFile Then
            'If no path then the document has not been saved
            If Not (myDoc.Path = "") Then
                'strNewFileName = Me.file_get_newFileName(myDoc, "wcag")
                strNewFullSavePath = Me.file_get_newFileName(myDoc)

                '
                'strNewFileName = Me.file_get_newFileName(myDoc, myDocInfo.DirectoryName, "wcag")
                '
                If objMsgMgr.msgMgr_dlg_doDocToWCAGExported() Then
                    '
                    'MsgBox(strNewFileName)
                    'Dim frm As New frm_Test()
                    'frm.txtBox_Path.Text = strNewFullSavePath
                    'frm.ShowDialog()
                    '
                    objGlobals.glb_cursors_setToWait()
                    '
                    Try
                        myDoc.Activate()
                        Me.file_get_saveTimeStampedCopy(myDoc, "wcag")
                        'myDoc.SaveAs2(strNewFullSavePath, Word.WdSaveFormat.wdFormatDocumentDefault)
                        '
                        'MsgBox($"Document saved as: {strNewFullSavePath}", MsgBoxStyle.Information)
                        '
                        'myDoc.Activate()
                        'myDoc.SaveAs2(strNewFullSavePath)
                        '
                        objWCAGMgr.wcag_doc_ToWCAG_Worker(myDoc)
                        '
                        'myDocInfo = New FileInfo(myDoc.FullName)
                        myDoc.SaveAs2()
                        '

                        'objWCAGMgr.wcag_doc_ToWCAG_entry(doTablesAsOutdented, myDoc)
                        'myDoc.Saved = False
                        'wrdApp.System.Cursor = WdCursorType.wdCursorNormal
                        '
                        '
                        objGlobals.glb_cursors_setToNormal()
                        '
                        MsgBox("The basic conversion is complete.." + vbCrLf + vbCrLf + "Your new partially 'Accessible' document can be found at" + vbCrLf + vbCrLf + strNewFullSavePath + vbCrLf + vbCrLf + "You'll need to engage 'Review > Check Accessibility' and adjust any remaining problems by hand")
                        'myDoc.Close()

                    Catch ex As Exception
                        'wrdApp.System.Cursor = WdCursorType.wdCursorNormal
                        MsgBox($"Conversion failed: {ex.Message}", MsgBoxStyle.Critical)

                        objGlobals.glb_cursors_setToNormal()
                    End Try
                    '
                    objGlobals.glb_cursors_setToNormal()
                    'newDoc.Close()

                    'wrdApp = Globals.ThisDocument.Application

                End If
            Else
                MsgBox("The document needs to be saved before it can be converted")
            End If
        Else
            'This is the inplace option
            'objWCAGMgr.wcag_doc_ToWCAG_entry(doTablesAsOutdented, myDoc)
            objWCAGMgr.wcag_doc_ToWCAG_Worker(myDoc)
            myDoc.Saved = False
            MsgBox("Please close, save and re-open the document to complete the 'Accessibility' conversion")

        End If
        '

    End Sub

    Public Sub file_doc_toWCAG(ByRef myDoc As Word.Document, doTablesAsOutdented As Boolean, Optional doAsNewFile As Boolean = True)
        Dim objMsgMgr As New cMessageManager()
        Dim objWCAGMgr As New cWCAGMgr()
        Dim myDocInfo As New FileInfo(myDoc.FullName)
        Dim strNewFileName As String
        '
        If doAsNewFile Then
            'new file option
            If myDocInfo.Exists Then
                strNewFileName = Me.file_get_newFileName(myDoc, myDocInfo.DirectoryName, "wcag")
                '
                If objMsgMgr.msgMgr_dlg_doDocToWCAGExported() Then
                    myDoc.SaveAs2(strNewFileName)
                    objWCAGMgr.wcag_doc_ToWCAG_entry(doTablesAsOutdented, myDoc)
                    myDoc.Saved = False
                    MsgBox("Please close, save and re-open the document to complete the 'Accessibility' conversion")
                    '
                End If
            Else
                MsgBox("The document needs to be saved before it can be converted")
            End If
        Else
            'This is the inplace option
            objWCAGMgr.wcag_doc_ToWCAG_entry(doTablesAsOutdented, myDoc)
            myDoc.Saved = False
            MsgBox("Please close, save and re-open the document to complete the 'Accessibility' conversion")

        End If
        '
    End Sub
    '
    '
    Public Function file_get_newFileName(strNewFileName As String, destinationFolderFullName As String, strFileId As String) As String
        Dim strExtension, strTimeSTamp As String
        Dim tokens() As String
        'Dim info As FileInfo
        '
        'info = My.Computer.FileSystem.GetFileInfo(strFullName)
        'strTimeSTamp = Me.getTimeStamp()
        strTimeSTamp = Me.file_get_TimeStamp()

        '
        'strExtension = Path.GetExtension(sourceDoc.FullName)
        tokens = strNewFileName.Split(".")
        strExtension = "." + tokens(1)
        'Now add an id to the file name (eg.g. wcag, dotNet etc)
        If strFileId = "" Then
            strNewFileName = tokens(0) + "-" + strTimeSTamp + strExtension
        Else
            strNewFileName = tokens(0) + "-" + strFileId + "-" + strTimeSTamp + strExtension
        End If
        '
        strNewFileName = destinationFolderFullName + strNewFileName
        '
        Return strNewFileName
    End Function
    '
    ''' <summary>
    ''' This method will return the local (and not the onedrive) version of the Documents directory
    ''' </summary>
    ''' <returns></returns>
    Public Function file_getDir_documentsLocal() As String
        Dim objGlobals As New cGlobals()
        Dim strLocalDocumentDir As String
        '
        strLocalDocumentDir = objGlobals.glb_getDir_documentsLocal()
        '
        Return strLocalDocumentDir
        '
    End Function

    '
    ''' <summary>
    ''' OneDrive safe.. Will save a timestamped copy of sourceDoc in the directory used
    ''' by sourceDoc (OneDrive or local). The return value is a string that is either null
    ''' 'Document saved as ' or 'Save failed '.. On exit sourceDoc is now the time stamped copy
    ''' </summary>
    ''' <param name="sourceDoc"></param>
    ''' <returns></returns>
    Public Function file_get_saveTimeStampedCopy(ByRef sourceDoc As Word.Document, Optional strFileId As String = "") As String
        Dim fullSavePath As String = ""
        Dim folderPath As String = ""
        Dim strMsg As String = ""
        Dim newFileName As String = System.IO.Path.GetFileNameWithoutExtension(sourceDoc.Name)
        '
        fullSavePath = Me.file_get_newFileName(sourceDoc, strFileId)

        Try
            sourceDoc.SaveAs2(fullSavePath, Word.WdSaveFormat.wdFormatDocumentDefault)
            'MsgBox($"Document saved as: {newFileName}", MsgBoxStyle.Information)
            strMsg = $"Document saved as: {newFileName}"
        Catch ex As Exception
            'MsgBox($"Save failed: {ex.Message}", MsgBoxStyle.Critical)
            strMsg = $"Save failed: {ex.Message}"
        End Try
        '
        Return strMsg
    End Function

    ''' <summary>
    ''' This is a OneDrive safe version that will return a time stamped filename derived from the sourceDoc
    ''' name aand Path. If strFileId is not null, then the strFileId is built into the string. If it
    ''' returns a null string then the Full path could not be developed... Probbably becuase the sourceDoc
    ''' has not been saved
    ''' </summary>
    ''' <param name="sourceDoc"></param>
    ''' <param name="strFileId"></param>
    ''' <returns></returns>
    Public Function file_get_newFileName(ByRef sourceDoc As Word.Document, Optional strFileId As String = "") As String
        Dim fullSavePath As String = ""
        Dim folderPath As String = ""
        Dim baseName As String = ""
        Dim timestamp As String = ""
        Dim newFileName As String = ""
        Dim strExtension As String = ""


        ' Ensure the document has a valid path
        If String.IsNullOrEmpty(sourceDoc.Path) Then
            fullSavePath = ""
        Else
            strExtension = Path.GetExtension(sourceDoc.FullName)

            ' Extract base name and timestamp
            baseName = System.IO.Path.GetFileNameWithoutExtension(sourceDoc.Name)
            timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss")
            '
            If strFileId = "" Then
                newFileName = $"{baseName}-{timestamp}{strExtension}"
            Else
                newFileName = $"{baseName}-{strFileId}-{timestamp}{strExtension}"
            End If
            '
            ' Build full URI or local path
            folderPath = sourceDoc.Path.TrimEnd("/"c, "\"c)
            fullSavePath = $"{folderPath}/{newFileName}"

        End If
        '
        Return fullSavePath
        '
    End Function
    '
    '
    ''' <summary>
    ''' Suitable for use on OneDrive... This function will provide a new timestamped filename
    ''' that can be saved back to OneDrive.. All prior versions were designed for local file system
    ''' </summary>
    ''' <param name="sourceDoc"></param>
    ''' <param name="strFileId"></param>
    ''' <returns></returns>
    Public Function file_get_newFileName_retired(ByRef sourceDoc As Word.Document, strFileId As String) As String
        Dim strExtension, strNewFileName, strTimeSTamp As String
        Dim tokens() As String
        'Dim info As FileInfo
        '
        strNewFileName = ""
        '
        If Not (sourceDoc.Path = "") Then
            '
            'info = My.Computer.FileSystem.GetFileInfo(strFullName)
            'strTimeSTamp = Me.getTimeStamp()
            strTimeSTamp = Me.file_get_TimeStamp()
            '
            strExtension = Path.GetExtension(sourceDoc.FullName)
            tokens = sourceDoc.Name.Split(".")
            'Now add an id to the file name (eg.g. wcag, dotNet etc)
            If strFileId = "" Then
                strNewFileName = tokens(0) + "-" + strTimeSTamp + strExtension
            Else
                strNewFileName = tokens(0) + "-" + strFileId + "-" + strTimeSTamp + strExtension
            End If
            '
            strNewFileName = sourceDoc.Path + "/" + strNewFileName
            '
        End If
        '
        Return strNewFileName
        '
    End Function


    '
    Public Function file_get_newFileName(ByRef sourceDoc As Word.Document, destinationFolderFullName As String, strFileId As String) As String
        Dim strExtension, strNewFileName, strTimeSTamp As String
        Dim tokens() As String
        'Dim info As FileInfo
        '
        'info = My.Computer.FileSystem.GetFileInfo(strFullName)
        'strTimeSTamp = Me.getTimeStamp()
        strTimeSTamp = Me.file_get_TimeStamp()
        '
        strExtension = Path.GetExtension(sourceDoc.FullName)
        tokens = sourceDoc.Name.Split(".")
        'Now add an id to the file name (eg.g. wcag, dotNet etc)
        If strFileId = "" Then
            strNewFileName = tokens(0) + "-" + strTimeSTamp + strExtension
        Else
            strNewFileName = tokens(0) + "-" + strFileId + "-" + strTimeSTamp + strExtension
        End If
        '
        strNewFileName = destinationFolderFullName + "\" + strNewFileName
        '
        Return strNewFileName
    End Function
    '


    '
    ''' <summary>
    ''' This method will take the source file as defined in sourceFile and chnage its name from sourceFile.xxx to
    ''' sourceFile_strFileId_yyyymmmdd.xxx
    ''' </summary>
    ''' <param name="sourceFile"></param>
    ''' <param name="destinationFolderFullName"></param>
    ''' <param name="strFileId"></param>
    ''' <returns></returns>
    Public Function file_get_newFileName(ByRef sourceFile As FileInfo, destinationFolderFullName As String, strFileId As String) As String
        Dim strExtension, strNewFileName, strTimeSTamp As String
        Dim tokens() As String
        'Dim info As FileInfo
        '
        'info = My.Computer.FileSystem.GetFileInfo(strFullName)
        strTimeSTamp = Me.file_get_TimeStamp()
        '
        strExtension = Path.GetExtension(sourceFile.FullName)
        tokens = sourceFile.Name.Split(".")
        'Now add an id to the file name (eg.g. wcag, dotNet etc)
        If strFileId = "" Then
            strNewFileName = tokens(0) + "-" + strTimeSTamp + strExtension
        Else
            strNewFileName = tokens(0) + "-" + strFileId + "-" + strTimeSTamp + strExtension
        End If
        '
        strNewFileName = destinationFolderFullName + "\" + strNewFileName
        '
        Return strNewFileName
    End Function
    '
    '
    ''' <summary>
    ''' Creates a timestamp string in a format that we'll use to append to files
    ''' This is typically used when saving/resaving multiple copies of a file over a 
    ''' period of time
    ''' </summary>
    ''' <returns></returns>
    Public Function file_get_TimeStamp() As String
        Dim objTimeStampMgr As New cTimeStampMgr()
        'Dim timeStamp As System.DateTime
        'Dim strTimeStamp As String

        'timeStamp = Now()
        'strTimeStamp = timeStamp.Year.ToString("D4") + timeStamp.Month.ToString("D2") + timeStamp.Day.ToString("D2") + "-" + timeStamp.Hour.ToString("D2") + timeStamp.Minute.ToString("D2") + timeStamp.Second.ToString("D2")
        file_get_TimeStamp = objTimeStampMgr.time_get_TimeStamp()
        '
    End Function
    '
    '
    ''' <summary>
    ''' Creates a timestamp string in a format that we'll use to append to files
    ''' This is typically used when saving/resaving multiple copies of a file over a 
    ''' period of time
    ''' </summary>
    ''' <returns></returns>
    Public Function getTimeStamp() As String
        Dim timeStamp As System.DateTime
        Dim strTimeStamp As String

        timeStamp = Now()
        strTimeStamp = timeStamp.Year.ToString("D4") + timeStamp.Month.ToString("D2") + timeStamp.Day.ToString("D2") + "-" + timeStamp.Hour.ToString("D2") + timeStamp.Minute.ToString("D2") + timeStamp.Second.ToString("D2")
        getTimeStamp = strTimeStamp
        '
    End Function
    '
    ''' <summary>
    ''' 
    ''' </summary>
    Public Sub xfile_get_templateFromWeb()
        Dim strWebPath As String
        Dim strFolderPath As String
        Dim strFileName As String
        Dim strFileFullPath As String
        Dim myFileInfo As FileInfo
        '
        'https://docs.microsoft.com/en-us/dotnet/visual-basic/developing-apps/programming/computer-resources/how-to-download-a-file
        '
        strFolderPath = "C:\Templates\"
        'strFileName = "GeneralReport_test.dotx"
        strFileName = Me.file_get_TimeStamp() + "_" + "AA GeneralReport.dotx"
        strFileFullPath = strFolderPath + strFileName
        strWebPath = "http://templates.acilallen.com.au/word/GeneralReport/install/AA GeneralReport.dotx"
        '
        Try
            My.Computer.Network.DownloadFile(strWebPath, strFileFullPath)
            'My.Computer.Network.DownloadFile(strWebPath, strFileFullPath, "username", "password")
            '
            'myFileInfo = New FileInfo("C:\Templates\GeneralReport.dotx")
            '
            Try
                'If myFileInfo.Exists Then
                'FileSystem.Rename("C:\Templates\GeneralReport.dotx", "C:\Templates\GeneralReport_old.dotx")
                'End If
            Catch ex1 As Exception

            End Try
            '
            '
            myFileInfo = New FileInfo("C:\Templates\GeneralReport-test.dotx")
            '
            Try
                If myFileInfo.Exists Then
                    'FileSystem.Rename("C:\Templates\GeneralReport-test.dotx", "C:\Templates\GeneralReport-test2.dotx")
                    'FileSystem.Rename("C:\Templates\GeneralReport-test.dotx", "C:\Templates\GeneralReport.dotx")
                End If
            Catch ex As Exception

            End Try
            '
            '
        Catch ex As Exception

        End Try
        '
        'If Not Directory.Exists(strFolderPath) Then
        'My.Computer.FileSystem.CreateDirectory(strFolderPath)
        'End If

        '
    End Sub
    '
    '
#Region "General Files"
    '
    ''' <summary>
    ''' This method will return true if the document has been saved.. or false if it has not
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function docSaveStatus(ByRef myDoc As Word.Document) As Boolean
        docSaveStatus = False
        '
        If myDoc.Path <> "" Then docSaveStatus = True
        '
    End Function
    '
    ''' <summary>
    ''' This method will determine of a file is accessible. Attempting to open a stream
    ''' with FileShare.None will throw a fault
    ''' </summary>
    ''' <param name="file"></param>
    ''' <returns></returns>
    Public Function isFileOpen(ByVal file As FileInfo) As Boolean
        Dim stream As FileStream
        '
        stream = Nothing
        isFileOpen = False
        Try
            stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None)
            stream.Close()
        Catch ex As Exception
            If TypeOf ex Is IOException Then
                isFileOpen = True
            End If
            Try
                stream.Close()
            Catch ex2 As Exception

            End Try
        End Try
        '
    End Function
    '

    Public Sub deleteFiles_OlderThan(lifeExpectancy As System.TimeSpan, strMyFolderPath As String)
        Dim lstOfFiles As Collections.ObjectModel.ReadOnlyCollection(Of String)
        Dim sb As New StringBuilder()
        Dim strFileName As String
        Dim myFileInfo As System.IO.FileInfo
        Dim dtNow_Local, createDate, dt As DateTime
        'Dim i As Integer
        '
        If My.Computer.FileSystem.DirectoryExists(strMyFolderPath) Then
            Try
                'Delete all files older or as old as 2 days
                dtNow_Local = DateTime.Now()
                '
                lstOfFiles = My.Computer.FileSystem.GetFiles(strMyFolderPath)
                '
                For Each strFileName In lstOfFiles
                    If My.Computer.FileSystem.FileExists(strFileName) Then
                        myFileInfo = My.Computer.FileSystem.GetFileInfo(strFileName)
                        createDate = myFileInfo.CreationTime()
                        '
                        dt = createDate.AddTicks(CDbl(lifeExpectancy.Ticks))
                        If dt <= dtNow_Local Then
                            My.Computer.FileSystem.DeleteFile(strFileName)
                        End If
                    End If
                Next
                '
            Catch ex1 As Exception
                'Something went wrong with the folder creation
                Exit Sub
            End Try
        Else
            'MessageBox.Show("'" + strMyFolderPath + "' " + "does not exist")
            'The directory does not exist so let's create it
            Try
                My.Computer.FileSystem.CreateDirectory(strMyFolderPath)
            Catch ex2 As Exception
                MessageBox.Show("Failed to Create " + strMyFolderPath + "You probably don't have the right permissions to do this." + vbCrLf + "Please contact your IT adminsitrator")
            End Try

        End If
    End Sub
    '
    ''' <summary>
    ''' This method will return a list of selected Word documents from the current directory
    ''' </summary>
    ''' <returns></returns>
    Public Function getDocsToImport(strDlgTitle As String, Optional ByVal strDlgFilter As String = "docx files (*.docx)|*.docx", Optional ByVal dlgMultiSelect As Boolean = True) As String()
        Dim strWorkingDirectory As String
        '
        strWorkingDirectory = My.Computer.FileSystem.CurrentDirectory
        '
        Return Me.file_get_filesFromDlg(strWorkingDirectory, strDlgTitle, strDlgFilter, dlgMultiSelect)
        '
    End Function

    ''' <summary>
    ''' This method will Return a list Of selected Word documents from the specified directory e.g. "C:\Templates."
    ''' You can specify the dialog box title, the file filter and whther it is multi select or not
    ''' </summary>
    ''' <param name="strInitialDirectoryName" purpose="The directory that the dialogbox will open"></param>
    ''' <param name="strDlgTitle" purpose="The dialog box title"></param>
    ''' <param name="strDlgFilter" purpose="The dialog box file filter string"></param>
    ''' <param name="dlgMultiSelect" purpose="True if multi select"></param>
    ''' <returns></returns>
    Public Function file_get_filesFromDlg(strInitialDirectoryName As String, strDlgTitle As String, Optional strDlgFilter As String = "docx files (*.docx)|*.docx", Optional ByVal dlgMultiSelect As Boolean = True) As String()
        Dim dlg As OpenFileDialog
        Dim strFileNames As String()
        Dim lstOfFiles As List(Of String)
        '
        Try
            lstOfFiles = New List(Of String)
            dlg = New OpenFileDialog()

            'dlg.Title = "The selected a Word documents will be exported as partially WCAG Compliant"
            dlg.Title = strDlgTitle
            dlg.InitialDirectory = strInitialDirectoryName

            'dlg.Filter = "docx files (*.docx)|*.docx"
            dlg.Filter = strDlgFilter
            dlg.FilterIndex = 1
            dlg.Multiselect = dlgMultiSelect
            dlg.RestoreDirectory = False
            dlg.CheckPathExists = True
            '
            strFileNames = Nothing
            '
            If dlg.ShowDialog() = DialogResult.OK Then
                strFileNames = dlg.FileNames()
                If strFileNames.Count > 0 Then
                    My.Computer.FileSystem.CurrentDirectory = My.Computer.FileSystem.GetParentPath(strFileNames(0))
                End If
            End If

        Catch ex As Exception
            'lstOfFiles = New List(Of String)
            strFileNames = Nothing
            Try
                My.Computer.FileSystem.CurrentDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Catch ex2 As Exception
                My.Computer.FileSystem.CurrentDirectory = "C:\"
            End Try
        End Try
        '
        Return strFileNames
    End Function


    '
#End Region

#Region "AutoSave Settings"
    Public Function getTimeIntervalStartIndex() As Integer
        Dim index As Integer
        Dim strIndex As String
        '
        strIndex = "0"
        Try
            'The line below will eventually retrieve a string value from the settings XML file. We
            'have two Converts just in case whatever is retrieved from the XML is corrupted
            strIndex = "2"
            index = Convert.ToInt32(strIndex)
        Catch ex As Exception
            strIndex = "0"
            index = Convert.ToInt32(strIndex)
            MessageBox.Show("Problem in retrieving Auto Save settings.. We'll be using the defaults")
        End Try

        Return index
    End Function
    '
    Public Function getNewFileName(ByRef sourceDoc As Word.Document, destinationFolderFullName As String, strFileId As String) As String
        Dim strExtension, strNewFileName, strTimeSTamp As String
        Dim tokens() As String
        'Dim info As FileInfo
        '
        'info = My.Computer.FileSystem.GetFileInfo(strFullName)
        'strTimeSTamp = Me.getTimeStamp()
        strTimeSTamp = Me.file_get_TimeStamp()
        '
        strExtension = Path.GetExtension(sourceDoc.FullName)
        tokens = sourceDoc.Name.Split(".")
        'Now add an id to the file name (eg.g. wcag, dotNet etc)
        If strFileId = "" Then
            strNewFileName = tokens(0) + "-" + strTimeSTamp + strExtension
        Else
            strNewFileName = tokens(0) + "-" + strFileId + "-" + strTimeSTamp + strExtension
        End If
        '
        strNewFileName = destinationFolderFullName + "\" + strNewFileName
        '
        Return strNewFileName
    End Function

    '
    ''' <summary>
    ''' This method will clone the sourceDoc to the new filename (which includes the full path)
    ''' as given by strNewFileName. The sourceDoc is left unaltered
    ''' </summary>
    ''' <param name="sourceDoc"></param>
    ''' <param name="strNewFileName"></param>
    ''' <returns></returns>
    Public Function cloneToSpecificPath(ByRef sourceDoc As Word.Document, strNewFileName As String) As Word.Document
        Dim newDoc As Word.Document
        Dim strExtension As String
        Dim wrdApp As Word.Application
        '
        newDoc = Nothing
        wrdApp = Me.objGlobals.glb_get_wrdApp
        strExtension = Path.GetExtension(sourceDoc.FullName)
        '
        Try
            Select Case strExtension
                Case ".docx", ".xml", ".txt"
                    newDoc = wrdApp.Documents.Add(sourceDoc.FullName, False,, False)          'Not visible
                    'newDoc = Globals.ThisDocument.Application.Documents.Add(sourceDoc.FullName, False,, True)           'Visible
                    '
                    newDoc.SaveAs2(strNewFileName, Word.WdSaveFormat.wdFormatDocumentDefault,,, False)
                    newDoc.Saved = True
                Case ".dotx", ".dotm"
            End Select
            '
        Catch ex As Exception
            'MsgBox("AutoSave only starts once the working file has been saved at least once")
            Try
                If Not IsNothing(newDoc) Then
                    newDoc.Close()
                End If
            Catch ex2 As Exception

            End Try
            newDoc = Nothing
        End Try
        '
finis:

        Return newDoc
    End Function

#End Region



End Class
