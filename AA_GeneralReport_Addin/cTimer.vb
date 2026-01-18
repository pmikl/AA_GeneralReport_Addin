Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports System.Diagnostics

Public Class cTimer
    Public Sub New()

    End Sub
    '
    Public Sub tmr_waitHere_milliseconds(ByRef IntervalinMs As Long)
        Dim endTime, startTime As Date
        Dim stpWatch As System.Diagnostics.Stopwatch
        '
        stpWatch = System.Diagnostics.Stopwatch.StartNew()
        startTime = TimeOfDay()
        '
loop00:
        '
        If stpWatch.ElapsedMilliseconds >= IntervalinMs Then
            stpWatch.Stop()
            endTime = TimeOfDay()
            GoTo finis
        Else
            If stpWatch.ElapsedMilliseconds >= 2000 Then
                MsgBox("Time sanity exit")
                GoTo finis
            End If
            GoTo loop00
        End If
        '
finis:

    End Sub

End Class
