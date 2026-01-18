Imports Microsoft.Win32
Public Class cGetDotNetVersion
    ' Calling the GetDotNetVersion.Get45PlusFromRegistry method produces 
    ' output like the following:
    '       .NET Framework Version: 4.6.1
    Public Sub New()

    End Sub
    Public Function Get45PlusFromRegistry() As String
        Dim strResult As String
        '
        strResult = ""
        Const subkey As String = "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\"
        Try
            Using ndpKey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(subkey)
                If ndpKey IsNot Nothing AndAlso ndpKey.GetValue("Release") IsNot Nothing Then
                    'Console.WriteLine(".NET Framework Version: " + CheckFor45PlusVersion(ndpKey.GetValue("Release")))
                    strResult = ".NET Framework Version: " + CheckFor45PlusVersion(ndpKey.GetValue("Release"))
                Else
                    'Console.WriteLine(".NET Framework Version 4.5 or later is not detected.")
                    strResult = ".NET Framework Version 4.5 or later is not detected."
                End If
            End Using
        Catch ex As Exception
            strResult = ""
        End Try
        '
        Return strResult
    End Function

    ' Checking the version using >= will enable forward compatibility.
    Private Shared Function CheckFor45PlusVersion(releaseKey As Integer) As String
        If releaseKey >= 460798 Then
            Return "4.7 or later"
        ElseIf releaseKey >= 394802 Then
            Return "4.6.2"
        ElseIf releaseKey >= 394254 Then
            Return "4.6.1"
        ElseIf releaseKey >= 393295 Then
            Return "4.6"
        ElseIf releaseKey >= 379893 Then
            Return "4.5.2"
        ElseIf releaseKey >= 378675 Then
            Return "4.5.1"
        ElseIf releaseKey >= 378389 Then
            Return "4.5"
        End If
        ' This code should never execute. A non-null release key should mean
        ' that 4.5 or later is installed.
        Return "No 4.5 or later version detected"
    End Function
End Class
