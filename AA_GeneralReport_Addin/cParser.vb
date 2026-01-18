Imports System
Imports System.Collections
Imports System.Text
Public Class cParser
    Private str2 As String
    '
    ''' <summary>
    ''' The parser must be initialised with a string on which to act
    ''' </summary>
    Public Property theString As String 'This is the angle theta
        Get
            Return Me.str2
        End Get
        Set(ByVal value As String) '
            Me.str2 = value
        End Set
    End Property
    '
    Public Sub New()
        Me.str2 = ""
    End Sub
    '
    Public Sub New(str1 As String)
        Me.str2 = str1
    End Sub
    '
    ''' <summary>
    ''' This method finds a token that ebgins with the string begin... If
    ''' strip is set to true, the begin string is removed from the token.
    ''' Otherwise it is left alone
    ''' </summary>
    ''' <param name="strip"></param>
    ''' <param name="begin"></param>
    ''' <param name="ListOfTokens"></param>
    ''' <returns></returns>
    Public Function findTokenBeginsWith(ByVal strip As Boolean, ByVal begin As String, ByVal ListOfTokens As ArrayList) As String
        Dim result = ""
        For Each Token As String In ListOfTokens
            If Token.Length >= begin.Length Then
                Dim st2 = Token.Substring(0, begin.Length)
                If Equals(st2, begin) Then
                    result = Token
                    Exit For
                End If
            End If
        Next
        If strip Then 'We must now strip away the begin string
            If result.Length = 0 Then Return result
            Dim str = result.Substring(begin.Length)
            result = str
        End If
        Return result
    End Function

    ''' <summary>
    ''' This routine returns an ArrayList of tokens that are extracted
    ''' from str2 (this value must be initialised). The toekn delimiter
    ''' is the char delim... It is not returned in the tokens
    ''' </summary>
    ''' <param name="delim"></param>
    ''' <returns></returns>
    Public Function getTokens(ByVal delim As Char) As ArrayList 'This is the single delimiter version
        Dim delimiters = New Char(0) {}

        delimiters(0) = delim
        Return Me.getTokens(delimiters)

    End Function

    Public Function getLines() As ArrayList
        Dim thechars = New Char(1) {}
        thechars(0) = ChrW(10)
        thechars(1) = ChrW(13)
        Return getTokens(thechars)
    End Function
    '
    ''' <summary>
    ''' This routine expects the object to have been initialised
    ''' with a value for str2. It returns an ArrayList that
    ''' contains a series of toekns (substrings). These tokens
    ''' are selected from str2 and are arbitray in
    ''' length. They are delimited by one or more characters that
    ''' are specified in the char[] delim. The delimiters are not
    ''' contained in the returned tokens
    ''' </summary>
    ''' <param name="delim"></param>
    ''' <returns></returns>
    Public Function getTokens(ByVal delim As Char()) As ArrayList
        Dim objal As ArrayList = New ArrayList()
        Dim isdelim As Boolean
        Dim sblen As Integer
        Dim tmp As Char() = str2.ToCharArray()
        Dim sb As StringBuilder = New StringBuilder(4)


        For i = 0 To tmp.Length - 1 'Lets go through each character
            isdelim = False
            'Is this character a delimeter
            For j = 0 To delim.Length - 1
                If tmp(i) = delim(j) Then
                    isdelim = True
                    Exit For
                End If
            Next
            Select Case isdelim
                Case True
                    'NO characters int eh string buffer
                    If sb.Length = 0 Then 'This is a delimeter and there are
                        'This is a delimeter and there are
                    Else
                        'characters int eh string buffer
                        objal.Add(sb.ToString())
                        sblen = sb.Length
                        sb.Remove(0, sblen)
                    End If
                Case False
                    'If we are at the end of the string, then we need
                    'to terminate with a new token
                    sb.Append(tmp(i).ToString())
                    If i = tmp.Length - 1 Then 'We are at the end.. but no delimiter
                        objal.Add(sb.ToString())
                        sblen = sb.Length
                        sb.Remove(0, sblen)
                    End If

                Case Else
            End Select

        Next
        Return objal
    End Function

End Class
