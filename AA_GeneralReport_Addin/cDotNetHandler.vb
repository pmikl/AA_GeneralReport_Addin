Imports System.IO
Imports System.Collections
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core

Public Class cDotNetHandler
    Inherits cGlobals
    '
    Public _strVBNetTemplate_TandG As String        'T and G version 2021
    Public _strVBNetTemplate As String              'Envelope version 2013
    Public _strVBATemplate As String

    Public _strConvertTag_dotNET_VBA_to_Env
    Public _strConvertTag_dotNET_Env_to_TandG
    Public _strConvertType_dotNET_VBA_to_TandG

    Public currentDoc As Word.Document
    Public _saveDirectoryFullName As String         'The directory into which the converted file is saved

    Public _strFileToConvert As String               'Full name of the file to convert
    Public _strFileTag As String
    Public _convertError As Boolean


    Public Sub New(strFileToConvertFullName As String, strSaveDirectory As String, fileTag As String)
        '
        Me._strVBNetTemplate_TandG = "C:\Templates\AA Report Template.dotx"
        Me._strVBNetTemplate = "C:\Templates\Acil Allen Template.dotx"
        Me._strVBATemplate = "Acil Allen Template.dotm"
        '
        Me._strConvertTag_dotNET_VBA_to_Env = "dotNET"
        Me._strConvertTag_dotNET_Env_to_TandG = "dotNET_Env_to_TandG"
        Me._strConvertType_dotNET_VBA_to_TandG = "dotNET_VBA_to_TandG"
        '
        Me.currentDoc = Nothing
        Me._strFileToConvert = strFileToConvertFullName
        Me._saveDirectoryFullName = strSaveDirectory
        Me._strFileTag = fileTag
        '
        Me._convertError = False

    End Sub
    '
    Public Sub New()
        Me._strConvertTag_dotNET_VBA_to_Env = "dotNET"
        Me._strConvertTag_dotNET_Env_to_TandG = "dotNET_Env_to_TandG"
        Me._strConvertType_dotNET_VBA_to_TandG = "dotNET_VBA_to_TandG"

    End Sub
    '


    Public Sub convertSelectedDoc()
        Dim newDoc As Word.Document
        Dim objFileMgr As New cFileHandler()
        Dim strNewFileName As String
        'Dim strCopyFullName As String
        Dim info As FileInfo
        'Dim evt As eve
        '
        newDoc = Nothing

        info = My.Computer.FileSystem.GetFileInfo(Me._strFileToConvert)
        '
        'Select Case strConversionType
        'Case Me._strConvertTag_dotNET_VBA_to_Env
        'Case Me._strConvertTag_dotNET_Env_to_TandG
        'Case Me._strConvertType_dotNET_VBA_to_TandG

        'End Select
        '
        Try
            If objFileMgr.isFileOpen(info) Then
                Me.currentDoc = Me.glb_get_wrdActiveDoc()
                'Me.currentDoc = Globals.ThisDocument.Application.Documents.Item(Me._strFileToConvert)
                strNewFileName = objFileMgr.getNewFileName(Me.currentDoc, Me._saveDirectoryFullName, Me._strFileTag)
                newDoc = Me.getEmptyNETDoc(strNewFileName)
                Me.modifyToDotNET(Me.currentDoc, newDoc)
                '
                newDoc.Save()
                newDoc.Close()
                Me._convertError = False
            Else
                'File is not open
                '
                'strCopyFullName = "tmp_" + info.Name
                'FileSystem.FileCopy(Me._strFileToConvert, strCopyFullName)
                Me.currentDoc = Me.glb_get_wrdApp.Documents.Open(FileName:=Me._strFileToConvert, AddToRecentFiles:=False, Visible:=False)
                'Me.currentDoc = Globals.ThisDocument.Application.Documents.Open(FileName:=strCopyFullName, AddToRecentFiles:=False, Visible:=False)

                strNewFileName = objFileMgr.getNewFileName(Me.currentDoc, Me._saveDirectoryFullName, Me._strFileTag)
                newDoc = Me.getEmptyNETDoc(strNewFileName)
                '
                'Me.modifyToDotNET(Me.currentDoc, newDoc)
                '
                Me.currentDoc.Saved = True
                Me.currentDoc.Close()
                'Me.modifyDocToWCAG(newDoc)
                '
                newDoc.Save()
                newDoc.Close()
                '
                Me._convertError = False
            End If
            '
        Catch ex As Exception
            Me._convertError = True
        End Try
    End Sub
    '
    Public Function modifyToDotNET(ByRef srcDocument As Word.Document, ByRef newNETDoc As Word.Document) As Word.Document
        Dim destinationDoc As Word.Document
        Dim rng, rngSource, rngDest As Word.Range
        '
        Me.glb_get_wrdApp.Options.PasteFormatBetweenStyledDocuments = WdPasteOptions.wdUseDestinationStyles
        destinationDoc = Nothing
        '
        Try
            newNETDoc.Activate()
            rngDest = newNETDoc.Sections(1).Range
            rngDest.Select()

            '
            For Each sect In srcDocument.Sections
                rngSource = sect.Range
                rngSource.Copy()
                rng = Me.glb_get_wrdApp.Selection.Range
                rng.PasteSpecial()                                          'Default Placement is WdOLEPlacement.wdFloatOverText
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                rng.Select()
            Next
            '
        Catch ex As Exception
            'targetDoc.Saved = True
            'targetDoc.Close()
            'copyContents = False
        End Try


        Return newNETDoc
    End Function
    '
    ''' <summary>
    ''' This method will rtrieve  a document that is based on the new VB.Net template 
    ''' 'Acil Allen Template.dotx'
    ''' </summary>
    ''' <returns></returns>
    Public Function getEmptyNETDoc(strNewFileName As String) As Word.Document
        Dim myDoc As Word.Document
        Dim strVBNetTemplate As String
        myDoc = Nothing
        '
        Select Case Me._strFileTag
            Case Me._strConvertTag_dotNET_VBA_to_Env
                strVBNetTemplate = Me._strVBNetTemplate
            Case Me._strConvertTag_dotNET_Env_to_TandG
                strVBNetTemplate = Me._strVBNetTemplate_TandG
            Case Me._strConvertType_dotNET_VBA_to_TandG
                strVBNetTemplate = Me._strVBNetTemplate_TandG
            Case Else
                strVBNetTemplate = Me._strVBNetTemplate_TandG

        End Select

        Try
            If My.Computer.FileSystem.FileExists(strVBNetTemplate) Then
                myDoc = Me.glb_get_wrdApp.Documents.Add(strVBNetTemplate, False,, False)
                'Save the document, but don't add it to the RecentFilesList... That's what the False is for
                'newDoc.SaveAs2(strFileName, Word.WdSaveFormat.wdFormatDocumentDefault,,, False)
                ' newDoc.SaveAs2(strFileName,,,, False)
                myDoc.SaveAs2(strNewFileName, Word.WdSaveFormat.wdFormatDocumentDefault,,, False)
                'newDoc.Saved = True
                'newDoc.Close()

            End If
        Catch ex As Exception

        End Try
        '
        Return myDoc
    End Function
    '
    ''' <summary>
    ''' This method will create the specified save directory if it does not exist. It will
    ''' set _createSaveDirectoriesIsOK = True on exit if the directory exists. it either existed
    ''' before entry, or was successfully created
    ''' </summary>
    Public Sub createSaveDirectory()
        If Not My.Computer.FileSystem.DirectoryExists(Me._saveDirectoryFullName) Then
            Try
                'Create the Directory
                My.Computer.FileSystem.CreateDirectory(Me._saveDirectoryFullName)
                'Me._createSaveDirectoriesIsOK = True
            Catch ex1 As Exception
                'Something went wrong with the folder creation
                'Me._createSaveDirectoriesIsOK = False
            End Try
        Else
            'Me._createSaveDirectoriesIsOK = True
        End If
    End Sub

End Class