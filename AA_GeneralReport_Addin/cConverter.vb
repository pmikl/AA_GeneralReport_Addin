Imports System
Imports System.Text
Imports System.Collections

Public Class cConverter
    Public Unicode As UnicodeEncoding = New UnicodeEncoding()
    Public Asciicode As ASCIIEncoding = New ASCIIEncoding()

    Public Sub New()

    End Sub
    '

    ''' <summary>
    ''' Converts a string to a byte array. Each string character
    ''' is treated as an ASCII code
    ''' </summary>
    ''' <param name="str1"></param>
    ''' <returns></returns>
    Public Function ConvertStringToByteArray(ByVal str1 As String) As Byte()
        Dim theBytes As Byte() = Asciicode.GetBytes(str1)
        Return theBytes


    End Function

    Public Sub ConvertUniCodeHexToString()


    End Sub

    ''' <summary>
    ''' This method expects a StringBuilder sbHexDigits to contains a
    ''' sequence of 2 character Hex digits separated by " ". The method will
    ''' return a StringBuilder that contains the String representation
    ''' of these Hex digits
    ''' </summary>
    ''' <param name="sbHexDigits"></param>
    ''' <returns></returns>
    Public Function ConvertASCIIHexToString(ByVal sbHexDigits As StringBuilder) As StringBuilder
        Dim theChars As Char()
        Dim sb As New StringBuilder()
        '
        theChars = Asciicode.GetChars(cConverter.convert_hexString_ToByteArray(sbHexDigits.ToString()))
        For i = 0 To theChars.Length - 1
            sb.Append(theChars(i))
        Next
        '
        Return sb
    End Function

    ''' <summary>
    ''' This routine takes the string str1 and returns a StringBuilder
    ''' containing the ASCII code (in Hex) for each character
    ''' </summary>
    ''' <param name="sbIn"></param>
    ''' <returns></returns>
    Public Function ConvertStringToAsciiHEX(ByVal sbIn As StringBuilder) As StringBuilder
        Dim i = 0

        Dim sb As StringBuilder = New StringBuilder()
        For i = 0 To sbIn.Length - 1

            sb.Append((Convert.ToByte(sbIn(i)).ToString("x").PadLeft(2, "0").ToUpper()))
            sb.Append(" ")
        Next

        Return sb
    End Function

    ''' <summary>
    ''' This routine takes the string str1 and returns a StringBuilder
    ''' containing the UniCode (in Hex) for each character
    ''' </summary>
    ''' <param name="str1"></param>
    ''' <returns></returns>
    Public Function ConvertStringToUniCodeHEX(ByVal str1 As String) As StringBuilder
        Dim sb As StringBuilder = New StringBuilder()
        For i = 0 To str1.Length - 1
            sb.Append((Convert.ToInt32(str1(i)).ToString("x").PadLeft(4, "0"c).ToUpper()))
            sb.Append(" "c)
        Next
        Return sb
    End Function




    ''' <summary>
    ''' This method expects as input a byte array. It will provide as output
    ''' a StrinBuilder containing a Hex respresentation (2 Hex digits) of
    ''' each byte. Each Hex number is separated by an ASCII space
    ''' </summary>
    ''' <param name="theBytes"></param>
    ''' <returns></returns>
    Public Shared Function ConvertByteArrayToHexString(ByRef theBytes As Byte()) As StringBuilder
        Dim sb As StringBuilder = New StringBuilder()
        For i = 0 To theBytes.Length - 1
            sb.Append(theBytes(i).ToString("x").PadLeft(2, "0").ToUpper())
            sb.Append(" ")
        Next

        Return sb
    End Function
    '
    ' 
    ''' <summary>
    ''' This method will take a string of 2 digit hex numbers and convert them to a byte array. If
    ''' something goes wrong the method will return nothing
    ''' </summary>
    ''' <param name="strHexDigits"></param>
    ''' <returns></returns>
    Public Shared Function convert_hexString_ToByteArray(strHexDigits As String) As Byte()
        ' remove any spaces from, e.g. "A0 20 34 34"
        strHexDigits = strHexDigits.Replace(" ", "")
        Dim nBytes = strHexDigits.Length \ 2
        Dim byteArray(nBytes - 1) As Byte
        '
        Try
            ' make sure we have an even number of digits
            If (strHexDigits.Length And 1) = 1 Then
                byteArray = Nothing
                GoTo finis
            End If

            ' calculate the length of the byte array and dim an array to that

            ' pick out every two bytes and convert them from hex representation
            For i = 0 To nBytes - 1
                byteArray(i) = Convert.ToByte(strHexDigits.Substring(i * 2, 2), 16)
            Next

        Catch ex As Exception
            byteArray = Nothing
        End Try

finis:
        Return byteArray

    End Function

End Class
