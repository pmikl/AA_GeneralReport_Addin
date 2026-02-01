Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports System.Windows.Forms
'
Public Class frm_WhatsNew
    Public objGlobals As cGlobals
    '
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        Me.objGlobals = New cGlobals()

        ' Add any initialization after the InitializeComponent() call.
        Me.whatsNew_text()
        Me.txtBox_Instruction.Text = "What's new information is available on youTube.. Just click on the latest youTube link below"
        'Me.txtBox_Instruction.Text = Me.txtBox_Instruction.Text + vbCrLf + vbCrLf + "The links are organised by date 'yyyymmdd'"

    End Sub
    '
    Public Sub whatsNew_text()
        Dim strText As String
        Dim textLink As New LinkLabel()
        'Dim Data As LinkLabel.Link

        strText = "7th Sept 2025:" + vbTab + vbTab + "Refined the fix to the 'Pages and Sections > Insert blank Lnd/Prt menu items'" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "- Refined the behaviour corrected in yesterday's update." + vbCrLf + vbCrLf
        '
        strText = strText + "6th Sept 2025:" + vbTab + vbTab + "Fix to the 'Pages and Sections > Insert blank Lnd/Prt menu items'" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "- Corrected the behaviour of the menu items. An error was introduced with the last'" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "  update. Text was unexpectedly deleted when any of the menu items were used." + vbCrLf + vbCrLf
        '
        strText = strText + "31st Aug 2025:" + vbTab + vbTab + "One solid change, one piece of groundwork" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "- Corrected the behaviour of the function under the menu item" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "  Finalise > Finalise > Individual document finishing Functions >" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "  1 space after end of all sentences." + vbCrLf + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "- Laid the groundwork for 1 second build times for Reports. " + vbCrLf + vbCrLf
        '
        strText = strText + "14th July 2025:" + vbTab + vbTab + "This set of changes concentrates on making it easier to insert sections (both " + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "landscape and portrait) into the ACIL Allen Brief bringing this new revision to" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "12.12.44/1.0.0.3758.... Check out the following 1920x1080 videos;" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "- 'Sections in the AA Brief' (5:25 min)" + vbTab + vbTab + vbTab + vbTab + vbTab + "https://youtu.be/MPV2K7ndRUs" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "- 'Word Fault - Sections sometimes disappear from view' (3:22 min)" + vbTab + vbTab + "https://youtu.be/GDGS4r6vzls" + vbCrLf + vbCrLf
        '
        strText = strText + "7th July 2025:" + vbTab + vbTab + "This is the first of a few small changes that will be made, bringing this release to " + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "12.12.44/1.0.0.3696." + vbCrLf + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "- In the body of the report, the Document Status WaterMark (e.g. 'DRAFT' and" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "  'DRAFT ONLY') is now located at the left side of the footer. This ensures that the " + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "  WaterMark is not covered by document elements such as the 'Key Findings' Placeholder." + vbCrLf + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "- In the next release the 'Brief' document type along with the 'bounded section parts'" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "  of the 'Insert blank Lnd/Prt Section' menu item will be modified to support easier" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "  insertion of bounded sections... In the meantime, if you want to insert a landscape" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "  section into a 'Brief' use 'Insert blank Lnd/Prt Section > Insert Blank Section at Selection'" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "  and then use 'Re-Orient to Lnd/Prt' to change the orientation of the new section." + vbCrLf + vbCrLf


        strText = strText + "24th Nov 2024:" + vbTab + vbTab + "As a result of user feedback, the following changes have been made," + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "bringing this release to 12.12.44/1.0.0.3604." + vbCrLf + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "- Changes to the Front Contacts Page  'Report to' and 'Proposal to' behaviour." + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "  The Copyright and Disclaimer statements can now be changed independent of" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "  the 'Report to' and 'Proposal to' sub headings." + vbCrLf + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "- Deployment over TLS (https) has been tested and verified." + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "  This is an internal 'housekeeping' issue." + vbCrLf + vbCrLf

        '
        strText = strText + "12th July 2024:" + vbTab + vbTab + "As a result of user feedback, the following changes have been made," + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "bringing this release to 12.12.44/1.0.0.3422." + vbCrLf + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "- For easier access 'Update TOC' placed in 'Pages and Sections > TOC Functions'." + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "- Standard tables are created with Heading Row repeat." + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "- Heading 1 functions now available for the 'Brief Report' type only." + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "- Option to insert 'bounded sections' with either standard or wide margins." + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "  See 'Pages and Sections > Insert blank Lnd/Prt Section'." + vbCrLf + vbCrLf

        strText = strText + "17th June 2024:" + vbTab + vbTab + "Changes to Styles, WaterMarks, Table Formatting and footers are now" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "going to 2 lines to cater for long report titles and sub titles. Have a look at the following " + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "1920 x 1080 videos..." + vbCrLf + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "'How to check which version of the template you are running' (2:54 min)" + vbTab + "https://youtu.be/YBIxNbfpHic" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "'How to get the most up to date Template and example documents' (2:54 min)" + vbTab + "https://youtu.be/Cfzkf8-HMJA" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "'Styles and Water Marks (8:30 min)'" + vbTab + vbTab + vbTab + vbTab + vbTab + "https://youtu.be/gtEYTjj08mU" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "'Rapid Table Format and the different table types (4:48 min)" + vbTab + vbTab + vbTab + "https://youtu.be/tts8rzPv34c" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "'Resetting your document to the new 2 line footer standard' (5:33 min)" + vbTab + vbTab + "https://youtu.be/C6Q87bACTmk" + vbCrLf + vbCrLf

        strText = strText + "2nd April 2024:" + vbTab + vbTab + "Too many changes to detail here. You'll have to refer to internal AA release notes." + vbCrLf + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "For some detail, have a look at the following 1920 x 1080  videos" + vbCrLf + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "'Creating Accessible documents' (9:04 min)" + vbTab + vbTab + vbTab + vbTab + "https://youtu.be/zlSpF1_MSQA" + vbCrLf + vbCrLf

        '
        strText = strText + "20th June 2022:" + vbTab + vbTab + "The AAC Ribbon now supports a new 'Front Contacts' page with options that includes; an 'Acknowledgment to Country'" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + " and an associated piece of art, or just standard text and no artwork. A 'new report build' defaults to the artwork version" + vbCrLf + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "The 'Accessibility (WCGA) Tools Group' has been moved to the 'Finalise' tab. It supports a converter that will convert a " + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "current 'well formed' AAC Report to an 'Accessible' document. Accessibility is determined by client requirements, so there" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "is likey to be more than one conversion option" + vbCrLf + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "Additional tools include 'Display Placeholders' and 'Ribbon Management'. Both are now available on the 'Finalise' tab either" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "in, or next to the 'Accessibility Tools Group'... The videos (below) on these are of slightly older versions. Whilst the" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "positions on the ribbon are slightly different, the functionality is the same." + vbCrLf + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "For more detail, have a look at the 1920 x 1080  videos" + vbCrLf + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "'Front Contacts Page and Accessibility Converter (16:07 min)" + vbTab + vbTab + vbTab + "https://youtu.be/7qx2cd9sDe4" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "'Ribbon Mgmnt' (4:09 min)" + vbTab + vbTab + vbTab + vbTab + vbTab + vbTab + "https://youtu.be/UNInPAjd5fk" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "'Display PlaceHolders' (4:18 min)" + vbTab + vbTab + vbTab + vbTab + vbTab + "https://youtu.be/aRK9bZOsR6c" + vbCrLf + vbCrLf

        '


        strText = strText + "29th May 2022:" + vbTab + vbTab + "The Accessibility (WCAG) work is still ongoing. The 'Convert to inline' function has been modified to" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + " what I think will be a more useful form. See the 1920x1080 video" + vbCrLf + vbCrLf
        '
        strText = strText + vbTab + vbTab + vbTab + "'Display PlaceHolders' (4:18 min)" + vbTab + vbTab + vbTab + vbTab + vbTab + "https://youtu.be/aRK9bZOsR6c" + vbCrLf + vbCrLf



        'Dim link = New LinkLabel()
        strText = strText + "23rd May 2022:" + vbTab + vbTab + "The Accessibility (WCAG) work is still ongoing. The main menu item for that has been shifted to the 'Finalise'" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "tab. When finished this will provide tools that will automatically convert an entire AAC report to as close to 'Accessible'" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "automatically convert an entire AAC report to as close to 'Accessible' as possible. Some tidying up " + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "may be required by the author." + vbCrLf + vbCrLf


        strText = strText + vbTab + vbTab + vbTab + "To support this work I did have to add some additional functions to the main tabs of the Report. Look at the " + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "following 1920x1080 videos;" + vbCrLf + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "'Ribbon Mgmnt' (4:09 min)" + vbTab + vbTab + vbTab + vbTab + vbTab + "https://youtu.be/UNInPAjd5fk" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "'Display PlaceHolders' (7:05 min)" + vbTab + vbTab + vbTab + "https://youtu.be/Y4ixOL0uFcQ" + vbCrLf + vbCrLf

        '
        '
        strText = strText + "10th April 2022:" + vbTab + vbTab + "The next two months will see a series of experimental" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "iterative changes being made (ONLY) to the WCAG" + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "Tools area of the Report." + vbCrLf + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "No other parts of the Report are being modified." + vbCrLf
        strText = strText + vbTab + vbTab + vbTab + "I will advise you here when the WCAG area is stable." + vbCrLf + vbCrLf


        strText = strText + "24th Oct 2021 (03:04 min):" + vbTab + "https://youtu.be/UsDNf8mLO7Q" + vbCrLf
        strText = strText + "10th Oct 2021 (08:40 min):" + vbTab + "https://youtu.be/PZjln_8Cq_Y"
        '
        Me.rTxtBx_WhatsNew.Text = Me.rTxtBx_WhatsNew.Text + strText


        'link.Text = "Microsoft"
        'Data = New LinkLabel.Link()
        'Data.LinkData =
        'link.Links.Add(Data)
        'link.Location = Me.rTxtBx_WhatsNew.GetPositionFromCharIndex(Me.rTxtBx_WhatsNew.TextLength)
        'Me.rTxtBx_WhatsNew.Controls.Add(link)
        '
        'yDoc = Me.rTxtBx_WhatsNew.do
        '
        '
        'ara = New Paragraph()
        'yDoc.b
        'e.rTxtBx_WhatsNew.te
    End Sub

    Private Sub rTxtBx_WhatsNew_LinkClicked(sender As Object, e As LinkClickedEventArgs) Handles rTxtBx_WhatsNew.LinkClicked
        'MsgBox("hello")
        System.Diagnostics.Process.Start(e.LinkText)
    End Sub

    Private Sub btn_Close_Click(sender As Object, e As EventArgs) Handles btn_Close.Click
        Me.Close()
    End Sub
End Class