Imports System.IO
''' <summary>
''' This class collects the function for handling the scratch directories for temporary
''' file dowloads (generally from Resources) and use. They are created and then deleted
''' </summary>
Public Class cFileScratchMgr
    Public _strScratchDir As String = "aa_scratch_wrd_addin"
    Public objGlobals As New cGlobals()
    Public Sub New()

    End Sub
    '
    ''' <summary>
    ''' This function returns the full path name of the scratch directory used for temporary
    ''' file downloads.. It is a subdirectory of the 'Templates' direcory
    ''' </summary>
    ''' <returns></returns>
    Public Function scratch_get_scratchDirectory() As String
        Dim strActualDirTemplates As String
        '
        strActualDirTemplates = objGlobals.glb_getDir_inUseforTemplates() + "\" + Me._strScratchDir
        '
        Return strActualDirTemplates
    End Function
    '
    '
    ''' <summary>
    ''' This method will build a local scratch directory. The directory is timestamped so that it's name
    ''' doesn't conflict with other directories
    ''' </summary>
    ''' <returns></returns>
    Public Function scratch_make_dirScratch() As String
        Dim strDirFullName As String
        '
        strDirFullName = Me.scratch_get_scratchDirectory()
        Me.scratch_make_dir(strDirFullName)
        '
        Return strDirFullName
        '
    End Function
    '
    ''' <summary>
    ''' This method will delete the 'scratch' directory
    ''' </summary>
    Public Sub scratch_delete_Directory_Scratch()
        Me.scratch_delete_Directory(Me.scratch_get_scratchDirectory())
    End Sub
    '
    ''' <summary>
    ''' This method will delete the specified directory and its contents
    ''' </summary>
    ''' <param name="strDirectoryFullName"></param>
    Private Sub scratch_delete_Directory(strDirectoryFullName As String)
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
    '
    Private Function scratch_make_dir(strDirectoryFullName As String) As Boolean
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


End Class
