Imports System.IO
Imports System.Windows.Forms
Public Class cFileHandlerOneDrive
    Public objGlobals As cGlobals
    Public Sub New()
        Me.objGlobals = New cGlobals()
    End Sub
    ''' <summary>
    ''' This method
    ''' </summary>
    ''' <param name="strFileType"></param>
    Public Sub oneDrv_get_OpenOneDriveFile(Optional strFileType As String = "all")
        Dim dialog As New OpenFileDialog()
        Dim oneDrivePath As String
        '
        'oneDrivePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "OneDrive")
        oneDrivePath = oneDrv_getDir_oneDrivePath()
        '
        If Not (oneDrivePath = "") Then
            If Directory.Exists(oneDrivePath) Then
                dialog.InitialDirectory = oneDrivePath
            Else
                'Make the initial directory the local Documents folder
                dialog.InitialDirectory = objGlobals.glb_getDir_documentsLocal()
            End If
        Else
            'Make the initial directory the local Documents folder
            dialog.InitialDirectory = objGlobals.glb_getDir_documentsLocal()
        End If
        '
        Select Case strFileType
            Case "all"
                dialog.Title = "Select a File"
                dialog.Filter = "All Files (*.*)|*.*"
            Case "word"
                dialog.Title = "Select a Word Document or Template"
                dialog.Filter = "Word Documents (*.docx)|*.docx|Word Templates (*.dotx)|*.dotx"
            Case Else
                dialog.Title = "Select a File"
                dialog.Filter = "All Files (*.*)|*.*"
        End Select
        '
        dialog.Multiselect = False
        '
        If dialog.ShowDialog() = DialogResult.OK Then
            Dim selectedFile As String = dialog.FileName
            'MessageBox.Show("Selected file: " & selectedFile)
            '
            Try
                ' Example: Open in Word
                Me.objGlobals.glb_get_wrdApp.Documents.Open(selectedFile)
                'Globals.ThisAddIn.Application.Documents.Open(selectedFile)

                ' Or for Excel:
                ' Globals.ThisAddIn.Application.Workbooks.Open(selectedFile)
            Catch ex As Exception

            End Try
        End If
    End Sub
    '
    ''' <summary>
    ''' This method will return the path to the OneDrive Shadow copies.. If it can't find them, then it returns an empty string
    ''' </summary>
    ''' <returns></returns>
    Public Function oneDrv_getDir_oneDrivePath() As String
        Dim oneDrivePath, oneDrivePersonal, oneDriveBusiness, oneDriveFallback As String
        '
        oneDrivePath = ""
        oneDrivePersonal = ""
        '
        Try
            'oneDrivePath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "OneDrive")
            'oneDrivePath = Environment.GetEnvironmentVariable("OneDriveConsumer")
            'oneDrivePath = Environment.GetEnvironmentVariable("OneDriveCommercial")
            '
            oneDrivePersonal = Environment.GetEnvironmentVariable("OneDriveConsumer")
            oneDriveBusiness = Environment.GetEnvironmentVariable("OneDriveCommercial")
            oneDriveFallback = Environment.GetEnvironmentVariable("OneDrive")
            '
            If Not IsNothing(oneDrivePersonal) Then
                oneDrivePath = oneDrivePersonal
            Else
                oneDrivePath = oneDriveFallback
            End If
            '
        Catch ex As Exception
            oneDrivePath = ""
        End Try
        '
        Return oneDrivePersonal
        '
    End Function
    '
    ''' <summary>
    ''' This method will return the local (and not the onedrive) version of the Documents directory
    ''' </summary>
    ''' <returns></returns>
    Public Function oneDrv_getDir_localDocuments() As String
        Dim objGlobals As New cGlobals()
        Dim strLocalDocumentDir As String
        '
        strLocalDocumentDir = objGlobals.glb_getDir_documentsLocal()
        '
        Return strLocalDocumentDir
        '
    End Function
    '
End Class
