Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cLegalAndAbout
    Public rgbPurpleText As Integer
    '
    Public Sub New()
        'Me.rgbPurpleText = RGB(108, 63, 153)
        Me.rgbPurpleText = RGB(157, 133, 190)
    End Sub
    '
    Public Function legal_insert_ackOfCountry(ByRef rng As Word.Range) As Word.Range
        Dim strToCountry As String
        '
        strToCountry = "ACIL Allen acknowledges Aboriginal and Torres Strait Islander peoples as the Traditional Custodians of the land and its waters. We pay our respects to Elders, past and present, and to the youth, for the future. We extend this to all Aboriginal and Torres Strait Islander peoples reading this report."
        '
        rng.Text = strToCountry
        '
        Return rng
    End Function
    '
    Public Function legal_insert_ackForArtWork(ByRef rng As Word.Range, strAck As String) As Word.Range
        '
        rng.Text = strAck
        '
        Return rng
    End Function
    '
    Public Function xxd() As String
        Dim strResult As String
        '
        strResult = "The following reliance and disclaimer is only for use in reports containing projections or modelling to support valuations or commercial strategy – please delete if not required and remember to delete this text" + vbCrLf
        '
        Return strResult
    End Function
    '
    Public Function insert_aboutACIlAllen(ByRef rng As Word.Range, ByRef myDoc As Word.Document) As Word.Range
        Dim strAbout As String
        '
        'strAbout = "ACIL Allen is the largest independent, Australian owned economic and public policy consultancy." + Chr(11)
        'strAbout = strAbout + "We specialise in the use of applied economics and econometrics with emphasis on the analysis, development and evaluation of policy, strategy and programs. Our reputation for quality research, credible analysis and innovative advice has been developed over a period of more than thirty years"
        '
        strAbout = "ACIL Allen is a leading independent economics, policy and strategy advisory firm, dedicated to helping clients solve complex issues." + vbCr
        strAbout = strAbout + "Our purpose is to help clients make informed decisions about complex economic and public policy issues." + vbCr
        strAbout = strAbout + "Our vision is to be Australia’s most trusted economics, policy and strategy advisory firm. We are committed and passionate about providing rigorous independent advice that contributes to a better world."

        rng.Text = strAbout


        Return rng
    End Function
    '
    Public Function insert_SuggestedCitation(ByRef rng As Word.Range, ByRef myDoc As Word.Document, Optional strDocType As String = "report_to") As Word.Range
        Dim strMsg As String
        Dim objGlobals As New cGlobals()
        Dim para As Word.Paragraph
        Dim rng2 As Word.Range
        Dim i As Integer
        '
        'strMsg = "Suggested citation for this report" + vbCrLf
        'strMsg = strMsg + "XXXXXXXXXX"
        '
        Select Case strDocType
            Case "report_to"
                strMsg = "Report to:" + vbCrLf
                strMsg = strMsg + "[Name of the organisation]"
            Case "proposal_to"
                strMsg = "Proposal to:" + vbCrLf
                strMsg = strMsg + "[Name of the organisation]"
            Case Else
                strMsg = "Report to:" + vbCrLf
                strMsg = strMsg + "[Name of the organisation]"

        End Select

        '
        rng.Text = strMsg
        '
        For i = 1 To rng.Paragraphs.Count
            para = rng.Paragraphs.Item(i)
            Select Case i
                Case 1, 2

                    para.Range.Style = objGlobals.glb_get_wrdActiveDoc.Styles.Item("Cp Disclaimer 9pt")
                    If i = 1 Then
                        rng2 = para.Range
                        'myStyle = rng2.Style
                        'myStyle.Font.Color = RGB(255, 0, 0)
                        'myStyle.Font.Size = 10
                        'myStyle.Font.Bold = True
                        rng2.MoveEnd(WdUnits.wdCharacter, -1)
                        'rng2.CharacterStyle.
                        rng2.Font.Size = 12
                        rng2.Font.Bold = True
                        rng2.Font.Color = objGlobals._glb_colour_purple_Dark

                    End If
                    If i = 2 Then
                        rng2 = para.Range
                        'myStyle = rng2.Style
                        'myStyle.Font.Color = RGB(255, 0, 0)
                        'myStyle.Font.Size = 10
                        'myStyle.Font.Bold = True
                        rng2.MoveEnd(WdUnits.wdCharacter, -1)
                        'rng2.CharacterStyle.
                        rng2.Font.Size = 11
                        rng2.Font.Bold = True
                        rng2.Font.Color = objGlobals._glb_colour_purple_Dark
                    End If
            End Select
        Next

        Return rng
    End Function
    '
    Public Function insert_disclaimer(ByRef rng As Word.Range, ByRef myDoc As Word.Document) As Word.Range
        Dim strResult As String
        Dim para As Word.Paragraph
        Dim rng2, rng3 As Word.Range
        Dim i As Integer
        '
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        strResult = ""
        '
        strResult = "The following reliance and disclaimer is only for use in reports containing projections or modelling to support valuations or commercial strategy – please delete if not required and remember to delete this text" + vbCrLf
        '
        'strResult = strResult + "Reliance and disclaimer The professional analysis and advice in this report has been prepared by ACIL Allen for the exclusive use of the party or parties to whom it Is addressed (the addressee) and for the purposes specified in it. This report Is supplied in good faith and reflects the knowledge, expertise and experience of the consultants involved. The report must not be published, quoted or disseminated to any other party without ACIL Allen's prior written consent. ACIL Allen accepts no responsibility whatsoever for any loss occasioned by any person acting or refraining from action as a result of reliance on the report, other than the addressee." + vbCrLf
        ' strResult = strResult + "In conducting the analysis in this report ACIL Allen has endeavoured to use what it considers is the best information available at the date of publication, including information supplied by the addressee. ACIL Allen has relied upon the information provided by the addressee and has not sought to verify the accuracy of the information supplied. Unless stated otherwise, ACIL Allen does not warrant the accuracy of any forecast or projection in the report. Although ACIL Allen exercises reasonable care when making forecasts or projections, factors in the process, such as future market behaviour, are inherently uncertain and cannot be forecast or projected reliably." + vbCrLf
        'strResult = strResult + "ACIL Allen shall not be liable in respect of any claim arising out of the failure of a client investment to perform to the advantage of the client or to the advantage of the client to the degree suggested or assumed in any advice or forecast given by ACIL Allen" + vbCrLf
        '
        strResult = strResult + "Reliance and disclaimer The professional analysis and advice in this report has been prepared by ACIL Allen for the exclusive use of the party or parties to whom it is addressed (the addressee) and for the purposes specified in it. This report is supplied in good faith and reflects the knowledge, expertise and experience of the consultants involved. The report must not be published, quoted or disseminated to any other party without ACIL Allen’s prior written consent. ACIL Allen accepts no responsibility whatsoever for any loss occasioned by any person acting or refraining from action as a result of reliance on the report, other than the addressee." + vbCrLf

        strResult = strResult + "In conducting the analysis in this report ACIL Allen has endeavoured to use what it considers is the best information available at the date of publication, including information supplied by the addressee. ACIL Allen has relied upon the information provided by the addressee and has not sought to verify the accuracy of the information supplied. If the information is subsequently determined to be false, inaccurate or incomplete then it is possible that our observations and conclusions as expressed in this report may change. The passage of time, manifestation of latent conditions or impacts of future events may require further examination of the project and subsequent data analysis, and re-evaluation of the data, findings, observations and conclusions expressed in this report.  Unless stated otherwise, ACIL Allen does not warrant the accuracy of any forecast or projection in the report. Although ACIL Allen exercises reasonable care when making forecasts or projections, factors in the process, such as future market behaviour, are inherently uncertain and cannot be forecast or projected reliably. ACIL Allen may from time to time utilise artificial intelligence (AI) tools in the performance of the services. ACIL Allen will not be liable to the addressee for loss consequential upon the use of AI tools." + vbCrLf
        strResult = strResult + "This report does not constitute a personal recommendation of ACIL Allen or take into account the particular investment objectives, financial situations, or needs of the addressee in relation to any transaction that the addressee is contemplating. Investors should consider whether the content of this report is suitable for their particular circumstances and, if appropriate, seek their own professional advice and carry out any further necessary investigations before deciding whether or not to proceed with a transaction. ACIL Allen shall not be liable in respect of any claim arising out of the failure of a client investment to perform to the advantage of the client or to the advantage of the client to the degree suggested or assumed in any advice or forecast given by ACIL Allen." + vbCrLf


        rng.Text = strResult
        '
        For i = 1 To rng.Paragraphs.Count
            para = rng.Paragraphs.Item(i)
            Select Case i
                Case 1
                    para.Style = myDoc.Styles.Item("Cp Disclaimer 9pt")
                    rng2 = para.Range
                    rng2.MoveEnd(WdUnits.wdCharacter, -1)
                    rng2.Shading.ForegroundPatternColor = RGB(0, 255, 255)
                Case 2
                    para.Style = myDoc.Styles.Item("Cp Disclaimer")
                    rng3 = para.Range
                    rng3.Collapse(WdCollapseDirection.wdCollapseStart)
                    rng3.MoveEnd(WdUnits.wdCharacter, 23)
                    rng3.Font.Bold = True
                Case 3, 4, 5
                    para.Style = myDoc.Styles.Item("Cp Disclaimer")
                    If i = 2 Then
                        'rng2 = para.Range
                        'rng2.Collapse(WdCollapseDirection.wdCollapseStart)
                        'rng2.MoveEnd(WdUnits.wdCharacter, 23)
                        'rng2.Font.Bold = True
                    End If
            End Select
            '
        Next
        '
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng.Fields.Add(rng, WdFieldType.wdFieldComments)
        '
        Return rng
    End Function

    '
    Public Function insert_CopyrightStatement(ByRef rng As Word.Range, ByRef myDoc As Word.Document) As Word.Range
        Dim strMsg As String
        Dim para As Word.Paragraph
        Dim i As Integer
        '
        Try
            strMsg = "Copyright in this document is and remains the property of ACIL Allen Pty Ltd. This document must not be reproduced in whole or in part without ACIL Allen's prior consent. Its content must only be used for the purposes of evaluation with a view to contracting ACIL Allen to carry out the work that is the subject matter of the document. No other use whatsoever can be made to any material or any recommendation, matter or thing in the document without ACIL Allen's prior written agreement" + vbCrLf
            rng.Text = strMsg
            '
            For i = 1 To rng.Paragraphs.Count
                para = rng.Paragraphs.Item(i)
                para.Range.Style = myDoc.Styles.Item("Cp Disclaimer")
            Next
            '
            '
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            rng.Fields.Add(rng, WdFieldType.wdFieldComments)
            '
        Catch ex As Exception
            MsgBox("Unable to insert 'Proposal Text'. Maybe a style is missing or has been renamed")
        End Try
        '
        Return rng
        '
    End Function
    '
    Public Function copyRightYear() As String
        Dim strMsg As String
        '
        strMsg = "© ACIL Allen 2024"
        '
        Return strMsg
        '
    End Function
    '
    '
    Public Sub insert_Back_MelbourneAndCanberra(ByRef rng As Word.Range, Optional numParasAfter As Integer = 0)
        Dim strMsg As String
        Dim para As Word.Paragraph
        Dim i As Integer
        Dim myDoc As Word.Document
        '
        myDoc = rng.Document
        '
        strMsg = "Melbourne" + vbCrLf
        strMsg = strMsg + "Suite 4, Level 19, North Tower" + vbCrLf
        strMsg = strMsg + "80 Collins Street" + vbCrLf
        strMsg = strMsg + "Melbourne VIC 3000 Australia" + vbCrLf
        strMsg = strMsg + "+61 3 8650 6000" + vbCrLf
        strMsg = strMsg + vbCrLf
        strMsg = strMsg + "Canberra" + vbCrLf
        strMsg = strMsg + "Level 6, 54 Marcus Clarke Street" + vbCrLf
        strMsg = strMsg + "Canberra ACT 2601 Australia" + vbCrLf
        strMsg = strMsg + "+61 2 6103 8200"
        '
        'For 2 paras after we need to add a vbcrlf after strMsg, then we must add 2 more. Therefore
        'for 2 paras after we need to run the loop 3 times
        If numParasAfter > 0 Then
            For i = 1 To numParasAfter + 1
                strMsg = strMsg + vbCrLf
            Next
        End If
        '
        rng.Text = strMsg
        '
        For i = 1 To rng.Paragraphs.Count
            para = rng.Paragraphs.Item(i)
            Select Case i
                Case 1, 7
                    para.Style = myDoc.Styles.Item("Cp Contact Details Heading")
                    'para.Range.Font.Color = Me.rgbPurpleText
                    'para.Range.Font.Bold = True
                Case Else
                    para.Style = myDoc.Styles.Item("Cp Contact Details")

            End Select
        Next
    End Sub
    '
    Public Sub insert_Back_SydneyAndPerth(ByRef rng As Word.Range, Optional numParasAfter As Integer = 0)
        Dim strMsg As String
        Dim para As Word.Paragraph
        Dim i As Integer
        Dim myDoc As Word.Document
        '
        myDoc = rng.Document
        '
        strMsg = "Sydney" + vbCrLf
        strMsg = strMsg + "Suite 603, Level 6" + vbCrLf
        strMsg = strMsg + "309 Kent Street" + vbCrLf
        strMsg = strMsg + "Sydney NSW 2000 Australia" + vbCrLf
        strMsg = strMsg + "+61 2 8272 5100" + vbCrLf
        strMsg = strMsg + vbCrLf
        strMsg = strMsg + "Perth" + vbCrLf
        strMsg = strMsg + "Level 12, 28 The Esplanade" + vbCrLf
        strMsg = strMsg + "Perth WA 6000 Australia" + vbCrLf
        strMsg = strMsg + "+61 8 9449 9600"
        '
        'For 2 paras after we need to add a vbcrlf after strMsg, then we must add 2 more. Therefore
        'for 2 paras after we need to run the loop 3 times
        '
        If numParasAfter > 0 Then
            For i = 1 To numParasAfter + 1
                strMsg = strMsg + vbCrLf
            Next
        End If
        '
        '
        rng.Text = strMsg
        '
        For i = 1 To rng.Paragraphs.Count
            para = rng.Paragraphs.Item(i)
            Select Case i
                Case 1, 7
                    para.Style = myDoc.Styles.Item("Cp Contact Details Heading")
                    'para.Range.Font.Color = Me.rgbPurpleText
                    'para.Range.Font.Bold = True
                Case Else
                    para.Style = myDoc.Styles.Item("Cp Contact Details")

            End Select
        Next
    End Sub
    '
    Public Sub insert_Back_BrisbaneAndAdelaide(ByRef rng As Word.Range, Optional numParasAfter As Integer = 0)
        Dim strMsg As String
        Dim para As Word.Paragraph
        Dim i As Integer
        Dim myDoc As Word.Document
        '
        myDoc = rng.Document
        '
        strMsg = "Brisbane" + vbCrLf
        strMsg = strMsg + "Level 15, 127 Creek Street" + vbCrLf
        strMsg = strMsg + "Brisbane QLD 4000 Australia" + vbCrLf
        strMsg = strMsg + "+61 7 3009 8700" + vbCrLf
        strMsg = strMsg + vbCrLf
        If numParasAfter = 0 Then strMsg = strMsg + vbCrLf
        strMsg = strMsg + "Adelaide" + vbCrLf
        strMsg = strMsg + "167 Flinders Street" + vbCrLf
        strMsg = strMsg + "Adelaide SA 5000 Australia" + vbCrLf
        strMsg = strMsg + "+61 8 8122 4965"
        '
        'For 2 paras after we need to add a vbcrlf after strMsg, then we must add 2 more. Therefore
        'for 2 paras after we need to run the loop 3 times
        '
        If numParasAfter > 0 Then
            For i = 1 To numParasAfter + 1
                strMsg = strMsg + vbCrLf
            Next
        End If
        '
        '
        rng.Text = strMsg
        '
        For i = 1 To rng.Paragraphs.Count
            para = rng.Paragraphs.Item(i)
            Select Case i
                Case 1, 7
                    para.Style = myDoc.Styles.Item("Cp Contact Details Heading")
                    'para.Range.Font.Color = Me.rgbPurpleText
                    'para.Range.Font.Bold = True
                Case Else
                    para.Style = myDoc.Styles.Item("Cp Contact Details")

            End Select
        Next
    End Sub
    '
    Public Sub insert_Back_CompanyAndABN(ByRef rng As Word.Range, Optional numParasAfter As Integer = 0)
        Dim strMsg As String
        Dim para As Word.Paragraph
        Dim i As Integer
        Dim myDoc As Word.Document
        '
        myDoc = rng.Document
        '
        strMsg = "ACIL Allen Pty Ltd" + vbCrLf
        strMsg = strMsg + "ABN 68 102 652 148"
        '
        '
        'For 2 paras after we need to add a vbcrlf after strMsg, then we must add 2 more. Therefore
        'for 2 paras after we need to run the loop 3 times
        '
        If numParasAfter > 0 Then
            For i = 1 To numParasAfter + 1
                strMsg = strMsg + vbCrLf
            Next
        End If
        '
        '
        rng.Text = strMsg
        '
        Try
            For i = 1 To rng.Paragraphs.Count
                para = rng.Paragraphs.Item(i)
                para.Style = myDoc.Styles.Item("Cp Contact Details")
            Next
        Catch ex As Exception

        End Try
        '
    End Sub
    '
    Public Sub insert_Back_WebAddress(ByRef rng As Word.Range, Optional numParasAfter As Integer = 0)
        Dim strMsg As String
        Dim para As Word.Paragraph
        Dim i As Integer
        Dim myDoc As Word.Document
        '
        myDoc = rng.Document
        '
        'rng.Style = rng.Document.Styles.Item("Cp Contact Details Web")
        strMsg = "acilallen.com.au"
        '
        'For 2 paras after we need to add a vbcrlf after strMsg, then we must add 2 more. Therefore
        'for 2 paras after we need to run the loop 3 times
        '
        If numParasAfter > 0 Then
            For i = 1 To numParasAfter + 1
                strMsg = strMsg + vbCrLf
            Next
        End If
        '
        rng.Text = strMsg
        '
        Try
            For i = 1 To rng.Paragraphs.Count
                para = rng.Paragraphs.Item(i)
                para.Style = myDoc.Styles.Item("Cp Contact Details Web")
            Next
        Catch ex As Exception

        End Try

        '
    End Sub
    '
    '
    Public Sub insert_CopyRightYear(myDoc As Word.Document)
        Dim objToolsMgr As New cTools()
        '
        Me.legal_upDate_CopyRightNotice()
        '
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will get the Copyright notice in the Comments fields of the Properties
    ''' area of the attached template and write it to the properties area of this document
    ''' I could save the document here, but for the moment we'll leave it be
    ''' </summary>
    Public Sub legal_upDate_CopyRightNotice()
        Dim objGlobals As New cGlobals()
        Dim strCommentsTmpl, strCommentThisDoc, strCopyRightField As String
        Dim strCopyrightStatement As String
        Dim builtInProperties As Microsoft.Office.Core.DocumentProperties
        Dim objFlds As New cFieldsMgr()
        'Dim objPropsMgr As New cPropertyMgr()
        Dim theDate As DateTime
        '
        theDate = DateTime.Now()

        strCopyRightField = "Comments"
        strCopyrightStatement = "© ACIL Allen " + Convert.ToString(Year(theDate))

        strCommentsTmpl = ""
        strCommentThisDoc = ""
        '
        'tmpl = objGlobals.glb_get_wrdActiveDoc.AttachedTemplate
        Try
            'strCommentsTmpl = strCopyrightStatement
            'Get the Coopyright Notice from the Attached Template
            ' strCommentsTmpl = CStr(tmpl.BuiltInDocumentProperties.Item(strCopyRightField).Value)
            'Somehow any manipulation of the template marks it as being chnaged. So, unless we
            'Set Saved to true the system will ask us whether we want to save any chnages made to the template.
            'Of course we can't have that
            'tmpl.Saved = True
            '
            'builtInProperties = Globals.ThisDocument.Application.ActiveDocument.AttachedTemplate.BuiltInDocumentProperties
            'strCommentsTmpl = CStr(Globals.ThisDocument.Application.ActiveDocument.AttachedTemplate.BuiltInDocumentProperties.Item(strCopyRightField).Value)
            'strCommentsTmpl = CStr(builtInProperties.Item(strCopyRightField).Value)
            '
            'objPropsMgr.setCustomProperty(strCommentsTmpl, "CopyRightNotice")
            '
            'Do Current Doc
            builtInProperties = objGlobals.glb_get_wrdActiveDoc.BuiltInDocumentProperties
            builtInProperties.Item(strCopyRightField).Value = strCopyrightStatement
            'Somehow any manipulation of the template marks it as being chnaged. So, unless we
            'Set Saved to true the system will ask us whether we want to save any chnages made to the template.
            'Of course we can't have that
            'tmpl.Saved = True
            '
            'Now update references to the Comments fields in the document
            objFlds.upDateCommentsField()
            '
        Catch ex As Exception
            builtInProperties = objGlobals.glb_get_wrdActiveDoc.BuiltInDocumentProperties
            builtInProperties.Item(strCopyRightField).Value = "© ACIL Allen "
            'Somehow any manipulation of the template marks it as being chnaged. So, unless we
            'Set Saved to true the system will ask us whether we want to save any chnages made to the template.
            'Of course we can't have that
            'tmpl.Saved = True
        End Try
        '
    End Sub
    '

End Class
