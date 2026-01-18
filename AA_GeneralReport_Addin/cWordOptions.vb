Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cWordOptions
    Inherits cGlobals
    Public Sub New()
        MyBase.New()
    End Sub
    '
    ''' <summary>
    ''' This method will set the paste picture to inline
    ''' </summary>
    Public Sub wrdOptions_set_pasteToInline()
        '
        If glb_get_wrdApp.Options.PictureWrapType <> WdWrapTypeMerged.wdWrapMergeInline Then
            glb_get_wrdApp.Options.PictureWrapType = WdWrapTypeMerged.wdWrapMergeInline
        End If
        '
    End Sub
    '
    ''' <summary>
    ''' This method will set the paste picture to inline
    ''' </summary>
    Public Sub wrdOptions_set_pasteInFrontOfText()
        '
        If glb_get_wrdApp.Options.PictureWrapType <> WdWrapTypeMerged.wdWrapMergeFront Then
            glb_get_wrdApp.Options.PictureWrapType = WdWrapTypeMerged.wdWrapMergeFront
        End If
        '
    End Sub
    '
    ''' <summary>
    ''' This method will set the field shading of the ActiveDocument depending on on the
    ''' value of strShadingMode. It can take on the values; 'always', 'never' or 'whenSelected'
    ''' <param name="strShadingMode"></param>
    ''' </summary>
    Public Sub wrdOptions_set_fieldShading(strShadingMode As String)
        glb_set_fieldShading(strShadingMode)
    End Sub

End Class
