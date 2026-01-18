Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
''' <summary>
''' 
'''Originally written in vba, some account taken for conversion
'''to vb.NET, but this was not a priority at this time
'''
'''Peter Mikelaitis April 2015...http://mikl.com.au
'''Ported to VB.NET 17th Jan 2017 from version 97p21p05
'''
''' </summary>
Public Class cBBlocksHandler
    Inherits cGlobals
    Public name As String
    Public tmpl As Word.Template
    '
    Public currentBBlockType As Integer
    '
    'Public defaultDocBuildingBlockLib As String             'Not used... will need to be deleted
    'Public defaultATextCategory As String                   'Default category for structural elements is defined here
    'Public defaultCorporateBuildingBlockLib As String       'Corporate Building Blocks are stored here
    '
    Public Sub New()
        '
        MyBase.New()

        Me.name = "hello"
        Me.tmpl = Globals.ThisAddIn.Application.ActiveDocument.AttachedTemplate
        '
        Me.currentBBlockType = WdBuildingBlockTypes.wdTypeCustom1
        '
    End Sub
    '
    Public Sub New(ByRef srcLib As Word.Template, strDefaultBBGallery As String)
        MyBase.New()

        Me.name = "hello"
        Me.tmpl = srcLib
        '
        Me.currentBBlockType = Me._getWdBuildingBlockType(strDefaultBBGallery)
        '
    End Sub
    '
    ''' <summary>
    ''' This function will take a string input and return the correct BuildingBlock (type/Gallery)
    ''' code. Alloed inputs are AutoText, Bibliography, CoverPage, Custom1, Custom2, Custom3
    ''' Custom4 and Custom5
    ''' </summary>
    ''' <param name="strBBlockType"></param>
    ''' <returns></returns>
    Public Function _getWdBuildingBlockType(strBBlockType As String) As Integer
        Dim rslt As Integer
        '
        rslt = WdBuildingBlockTypes.wdTypeAutoText
        Select Case strBBlockType
            Case "AutoText"
                rslt = WdBuildingBlockTypes.wdTypeAutoText
            Case "Bibliography"
                rslt = WdBuildingBlockTypes.wdTypeBibliography
            Case "CoverPage"
                rslt = WdBuildingBlockTypes.wdTypeCoverPage
            Case "Custom1"
                rslt = WdBuildingBlockTypes.wdTypeCustom1
            Case "Custom2"
                rslt = WdBuildingBlockTypes.wdTypeCustom2
            Case "Custom3"
                rslt = WdBuildingBlockTypes.wdTypeCustom3
            Case "Custom4"
                rslt = WdBuildingBlockTypes.wdTypeCustom4
            Case "Custom5"
                rslt = WdBuildingBlockTypes.wdTypeCustom5
        End Select
        '
        _getWdBuildingBlockType = rslt
    End Function


    '
    '
    'This method will determine whether the current selection is OK for the
    'insert section method
    Public Function canInsert(strType As String) As Boolean
        Dim sect As Word.Section
        '
        If strType Like "_Landscape" Then

        End If

        sect = Globals.ThisAddIn.Application.Selection.Sections(1)                                                'Set the first section..normally the only section
        canInsert = True
        '
        If Globals.ThisAddIn.Application.Selection.Tables.Count > 0 Then canInsert = False            'If in a Table don't allow it
        If Globals.ThisAddIn.Application.Selection.Sections.Count > 1 Then canInsert = False          'If spans sections then don't allow it
        If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape _
        Then canInsert = False                                      'Do not allow insertion into existing Landscape section
        If sect.Index = 1 Then canInsert = False                        'If in title section, then don't allow it
        If sect.Index = 2 Then canInsert = False                        'If in disclaimer section, then don't allow it
        If sect.Index = 3 Then canInsert = False                        'If in TOC section, then don't allow it

    End Function
    '
    'This method will obtain the building block with the name strType from
    'the default document Building Block type library and insert it at the
    'current cursor position. This method will return the insert Range
    '
    Public Function insertBuildingBlockFromDefaultLib_ReturnRange(strBBName As String, strCategoryName As String) As Range
        Dim objBB As Word.BuildingBlock
        Dim rng As Word.Range
        Dim strMsg As String
        '
        strMsg = "The building block called " & strBBName & " could not be inserted"
        insertBuildingBlockFromDefaultLib_ReturnRange = Nothing
        '
        Try
            objBB = Me.getBuildingBlockFromDefaultLib(strBBName, strCategoryName)
            rng = Globals.ThisAddIn.Application.Selection.Range
            Call objBB.Insert(rng, True)
            insertBuildingBlockFromDefaultLib_ReturnRange = rng
        Catch ex As Exception
            MsgBox(strMsg)
        End Try
        '
    End Function
    '
    ''' <summary>
    ''' This method will insert the specified building block (obtained from Me.currentBlockType..defaults to Custom1)
    ''' at the Range destRng. It will return the range of the inserted Building Block
    ''' </summary>
    ''' <param name="destRng"></param>
    ''' <param name="strBBName"></param>
    ''' <param name="strCategoryName"></param>
    ''' <returns></returns>
    Public Function _insertBuildingBlockToRange(destRng As Word.Range, strBBName As String, strCategoryName As String) As Range
        Dim objBB As BuildingBlock
        Dim strMsg As String
        '
        strMsg = "The building block called " & strBBName & " could not be inserted"
        Try
            objBB = Me.getBuildingBlockFromDefaultLib(strBBName, strCategoryName)
            _insertBuildingBlockToRange = objBB.Insert(destRng, True)
        Catch ex As Exception
            _insertBuildingBlockToRange = Nothing
            MsgBox(strMsg)
        End Try
        '
    End Function
    '
    ''' <summary>
    ''' This function accesses the Custom1 BuildingBlock type (i.e.lib) and inserts, at the
    ''' current Selection, the building block with the name strBBName from the Category strCategoryName.
    ''' It returns a range object that represents the contents of the building block within the document.
    ''' Example (from Custom1);
    ''' strBBName = 'contactsFront'
    ''' strCategoryName = 'CoverPage'
    ''' </summary>
    ''' <param name="strBBName"></param>
    ''' <param name="strCategoryName"></param>
    ''' <returns></returns>
    Public Function _insertBBlockToSelection(strBBName As String, strCategoryName As String) As Word.Range
        Dim objBBlk As Word.BuildingBlock
        Dim rng, oldRange As Word.Range
        Dim strMsg As String
        '
        strMsg = "The building block called " & strBBName & " could not be inserted"
        '
        oldRange = Globals.ThisAddIn.Application.Selection.Range
        '
        Try
            rng = Globals.ThisAddIn.Application.Selection.Range
            'objBBlk = Me.tmpl.BuildingBlockTypes.Item(WdBuildingBlockTypes.wdTypeCustom1).Categories.Item("CoverPage").BuildingBlocks.Item("contactsFront")
            objBBlk = Me.tmpl.BuildingBlockTypes.Item(Me.currentBBlockType).Categories.Item(strCategoryName).BuildingBlocks.Item(strBBName)
            rng = objBBlk.Insert(rng)
            Return rng
        Catch ex As Exception
            MsgBox(strMsg)
            Return oldRange
        End Try

    End Function
    '
    'This method will obtain the building block with the name strType from
    'the default document Building Block type library and insert it at the
    'current cursor position. This method will return the Range of the
    'Building Block
    '
    Public Function insertBuildingBlockFromDefaultLib_ReturnRangeOfBlock(strBBName As String, strCategoryName As String) As Range
        'Dim objBB As BuildingBlock
        Dim rng As Range
        'Dim strMsg As String
        '
        '*** Test replacement
        rng = Me._insertBBlockToSelection(strBBName, strCategoryName)
        '*** 
        '
        'strMsg = "The building block called " & strBBName & " could not be inserted"
        'On Error GoTo finis
        '
        'objBB = Me.getBuildingBlockFromDefaultLib(strBBName, strCategoryName)
        'rng = Globals.ThisAddin.Application.Selection.Range
        'rng = objBB.Insert(rng, True)
        'insertBuildingBlockFromDefaultLib_ReturnRangeOfBlock = rng
        '
        'Exit Function
        'finis:
        'MsgBox(strMsg)
        '
        Return rng
    End Function
    '
    '
    'This method will obtain the building block with the name strType from
    'the default document Building Block type library and insert it at the
    'current cursor position.. It has no return
    Public Function insertBuildingBlockFromDefaultLib(strBBName As String, strCategoryName As String) As Word.Range
        Dim rng As Range
        '
        rng = Me.insertBuildingBlockFromDefaultLib_ReturnRange(strBBName, strCategoryName)
        '
        Return rng
    End Function
    '
    'This method will obtain the building block with the name strType from
    'the default document Building Block type library and insert
    'it at the specified range
    '
    Public Function insertBuildingBlockFromDefaultLibToRange(strBBName As String, strCategoryName As String, ByRef rng As Range) As Range
        Dim objBB As BuildingBlock
        Dim objGlobals As New cGlobals()
        Dim strMsg As String
        Dim rngInsert As Word.Range
        Dim myDoc As Word.Document
        '
        myDoc = rng.Document
        'myDoc.AttachedTemplate = objGlobals.glb_var_TemplateFileName()
        'strTemplateFullName = objGlobals.glb_getTmpl_FullName()
        '
        'myDoc.AttachedTemplate = objGlobals.glb_getTmpl_FullName()


        '
        strMsg = "The building block called " & strBBName & " could not be inserted"

        Try
            objBB = Me.getBuildingBlockFromDefaultLib(strBBName, strCategoryName)
            rngInsert = objBB.Insert(rng, True)
            '
        Catch ex As Exception
            rngInsert = Nothing
            'MsgBox(strMsg)
        End Try
        '
        'myDoc.AttachedTemplate = "Normal"
        '
        Return rngInsert
        '
    End Function
    '
    'This method will retrieve a Building Block from the default corporate building block
    'type library as identified by Me.defaultCorporateBuildingBlockLib;
    'strBBName:         Building Block Name
    'strCategoryName:   The category in which the building block resides
    '
    Public Function getBuildingBlockFromDefaultLib(strBBName As String, strCategoryName As String) As Word.BuildingBlock
        '
        Try
            getBuildingBlockFromDefaultLib = Me.tmpl.BuildingBlockTypes.Item(Me.currentBBlockType).Categories.Item(strCategoryName).BuildingBlocks.Item(strBBName)
        Catch ex As Exception
            getBuildingBlockFromDefaultLib = Nothing
        End Try
        'getBuildingBlockFromDefaultLib = Me.getBuildingBlockEntry(strBBName, Me.defaultCorporateBuildingBlockLib, strCategoryName)
    End Function

    '
    Public Function getBuildingBlockEntry(strBBName As String, strCategoryName As String, wdBuildingBlockType As Integer) As BuildingBlock
        Try
            getBuildingBlockEntry = Me.tmpl.BuildingBlockTypes.Item(wdBuildingBlockType).Categories.Item(strCategoryName).BuildingBlocks.Item(strBBName)
        Catch ex As Exception
            getBuildingBlockEntry = Nothing
        End Try
    End Function
    '
    Public Function bblk_insert_sectionInFront(Optional strSectionType As String = "aa_Chpt_Std") As Range
        Dim objDivMgr As New cChptDivider()
        Dim objTblsMgr As New cTablesMgr()
        Dim sect As Word.Section
        Dim para As Word.Paragraph
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim myDoc As Word.Document
        '
        sect = glb_get_wrdSect()
        rng = sect.Range
        myDoc = rng.Document
        '
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()
        '
        If glb_selection_IsInTable() Then
            'The section has a table at the top of the section. So, to insert a new section building block
            'we need to get the table, place a paragraph above it, then select the returned (collapsed) range
            'to ensure the insertion goes as expected
            tbl = glb_get_wrdSelTbl()
            rng = objTblsMgr.tbl_para_addAbove2(tbl)
            para = rng.Paragraphs.First()
            rng.Select()
            '
            'Now insert the portrait or landscape building block, depending 
            Select Case sect.PageSetup.Orientation
                Case WdOrientation.wdOrientPortrait
                    rng = Me.insertBuildingBlockFromDefaultLib_ReturnRangeOfBlock(strSectionType + "_Prt", "aa_ReportPrt")
                Case WdOrientation.wdOrientLandscape
                    rng = Me.insertBuildingBlockFromDefaultLib_ReturnRangeOfBlock(strSectionType + "_Lnd", "aa_ReportLnd")
            End Select
            rng = objDivMgr.chptBase_getRange_Heading1(rng.Sections.First)
            rng.Select()
            '
            para.Range.Delete()
        Else
            'The selection is not ina table, but it is at the beginning of a section
            Select Case sect.PageSetup.Orientation
                Case WdOrientation.wdOrientPortrait
                    rng = Me.insertBuildingBlockFromDefaultLib_ReturnRangeOfBlock(strSectionType + "_Prt", "aa_ReportPrt")
                Case WdOrientation.wdOrientLandscape
                    rng = Me.insertBuildingBlockFromDefaultLib_ReturnRangeOfBlock(strSectionType + "_Lnd", "aa_ReportLnd")
            End Select
            rng = objDivMgr.chptBase_getRange_Heading1(rng.Sections.First)
            rng.Select()


        End If
        '
finis:
        Return rng
    End Function
End Class
