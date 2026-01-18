Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports System.Drawing
''' <summary>
''' This class deals with acess to and manipulation of data from the local Resources
''' file
''' </summary>
Public Class cResourcesMgr
    Public Sub New()

    End Sub
    '
    ''' <summary>
    ''' This method gets the image used in the Chapter banners
    ''' </summary>
    ''' <returns></returns>
    Public Function rsrcs_get_bannerImage() As Image
        Dim img As Image
        '
        img = Nothing
        '*** Insert from Resources
        'img = My.Resources.aac_Img_chptBanner_std
        'Me.Banner_Std_Image = My.Resources.banner_KI_Sunset_03
        'Me.Banner_Std_Image = My.Resources.banner_Nasa_bg_image
        '
        Return img
    End Function
    '
    ''' <summary>
    ''' This method gets the artwork used in the bottom area of the front Contacts
    ''' page
    ''' </summary>
    ''' <returns></returns>
    Public Function rsrcs_get_contactsArtWork() As Image
        Dim img As Image
        '
        Try
            img = My.Resources.artwork_contactsPage_front_release
        Catch ex As Exception
            img = Nothing
        End Try
        '
        Return img
        '
    End Function

End Class
