Public Class cTimeStampMgr
    Inherits cFileHandler
    Public Sub New()
        MyBase.New()
    End Sub
    '
    '
    ''' <summary>
    ''' Creates a timestamp string in a format that we'll use to append to files
    ''' This is typically used when saving/resaving multiple copies of a file over a 
    ''' period of time
    ''' </summary>
    ''' <returns></returns>
    Public Function time_get_TimeStamp() As String
        Dim strTimeStamp As String

        strTimeStamp = DateTime.Now.ToString("yyyyMMdd-HHmmss")

        'timeStamp = Now()
        'strTimeStamp = timeStamp.Year.ToString("D4") + timeStamp.Month.ToString("D2") + timeStamp.Day.ToString("D2") + "-" + timeStamp.Hour.ToString("D2") + timeStamp.Minute.ToString("D2") + timeStamp.Second.ToString("D2")
        '
        Return strTimeStamp
    End Function
    '
    ''' <summary>
    ''' This method will take the file name (not full apth) as defined in strNewFileName and modify it from strNewFileName to
    ''' strNewFileName_strFileId_yyyymmmdd.xxx. It will return a string which is a fullPathName (includes the 
    ''' destinationFolderFullName)... This is typically used when you want to resave a document under a new
    ''' name. In this case the new name includes a time stamp
    ''' </summary>
    ''' <param name="strNewFileName"></param>
    ''' <param name="destinationFolderFullName"></param>
    ''' <param name="strFileId"></param>
    ''' <returns></returns>
    Public Function time_get_newFileName(strNewFileName As String, destinationFolderFullName As String, strFileId As String) As String
        Dim strExtension, strTimeSTamp As String
        Dim tokens() As String
        'Dim info As FileInfo
        '
        'info = My.Computer.FileSystem.GetFileInfo(strFullName)
        strTimeSTamp = Me.time_get_TimeStamp()
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
    ''' This method will take the source file as defined in sourceFile and chnage its name from sourceFile.xxx to
    ''' sourceFile_strFileId_yyyymmmdd.xxx. It will return a string which is a fullPathName (includes the 
    ''' destinationFolderFullName)... This is typically used when you wnat to resave a document under a new
    ''' name name. In this case the new name includes a time stamp
    ''' </summary>
    ''' <param name="sourceDoc"></param>
    ''' <param name="destinationFolderFullName"></param>
    ''' <param name="strFileId"></param>
    ''' <returns></returns>
    Public Function time_get_newFileName(ByRef sourceDoc As Word.Document, destinationFolderFullName As String, strFileId As String) As String
        Dim strExtension, strNewFileName, strTimeSTamp As String
        Dim tokens() As String
        'Dim info As FileInfo
        '
        'info = My.Computer.FileSystem.GetFileInfo(strFullName)
        strTimeSTamp = Me.time_get_TimeStamp()
        '
        strExtension = System.IO.Path.GetExtension(sourceDoc.FullName)
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
    ''' sourceFile_strFileId_yyyymmmdd.xxx. It will return a string which is a fullPathName (includes the 
    ''' destinationFolderFullName)... This is typically used when you wnat to resave a document under a new
    ''' name name. In this case the new name includes a time stamp
    ''' </summary>
    ''' <param name="sourceFile"></param>
    ''' <param name="destinationFolderFullName"></param>
    ''' <param name="strFileId"></param>
    ''' <returns></returns>
    Public Function time_get_newFileName(ByRef sourceFile As System.IO.FileInfo, destinationFolderFullName As String, strFileId As String) As String
        Dim strExtension, strNewFileName, strTimeSTamp As String
        Dim tokens() As String
        'Dim info As FileInfo
        '
        'info = My.Computer.FileSystem.GetFileInfo(strFullName)
        strTimeSTamp = Me.time_get_TimeStamp()
        '
        strExtension = System.IO.Path.GetExtension(sourceFile.FullName)
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
End Class
