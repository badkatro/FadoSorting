Public strKeys() As String      ' string sorting keys array



'
' **********************************************************************************************************
' ***********   Routines for sorting blocks of text, made for ST 6781/12 REV1 & ST 9916/08 *****************
' ***********   That means for                  "Glosarul FADO"                             ****************
' **********************************************************************************************************
'

Sub Master_Process_Doc_For_Alignment()
' Also does sorting of "chapters" (identified by a header with a one row table containing word "top" and an arrow)

'Call CoverNote_Remove
' Bookmarks will be used for sorting, we need document to be clean for that
Call All_Bookmarks_Remove(ActiveDocument)
' All little blue arrows are out, as well as most pictures (the ones from tables at least)
Call AllInlineImagesDelete
' All identifying "top" tables are highlighted (shaded in fact), rest of tables
' are converted to text
Call Highlight_All_TopTables_CvTxt_Rest
' Clean empty enters
Call AllVBCr_Pairs_ToVbCr
Call AllMultiple_VBCrs_To_Single
Call AllPageBreaks_ToEnters


End Sub


Sub Master_Sort_ActiveDoc_FadoGlossary()

' Sanity check
If MsgBox("Have you already deleted the LAST top table from this document", vbYesNo + vbQuestion, "Check carefully!") = vbNo Then
    StatusBar = "Please delete it and retry..."
    Exit Sub
End If

Dim nrTopTables As Integer

' count top tables IF THEY'RE still signalled by shading color index 16 !!! CHECK!
nrTopTables = Count_TopTables(ActiveDocument)

' set top bkm before each chapter
Call Set_Top_Bookmarks(ActiveDocument, False)

' create containing bookmarks for each chapter (these will be moved arount at sorting)
Call Set_MainBookmarks(ActiveDocument, nrTopTables)

' now extract sorting keys (first non numeric paragraphs from each container bookmark)
Call Create_Fado_SortingKeys_ForAllBookmarks("Fado_", nrTopTables)

' show user info form
frmSortPages.Show

' and start sorting
Call Bkm_Sort("Fado_", nrTopTables)

' and clean-up
Unload frmSortPages


End Sub


Sub Master_Sort_TargetDoc_forAlignment_FadoGlossary()
' Sorts target document chapters acc to original document ids order
' Useful so as to align 2 FADOs documents, since they need be in same order for alignment success

' TO BE RAN from the target document, source document opened as well !!


If Documents.Count <> 2 Then
    MsgBox "Please open the source and target documents, last top table removed, no cover page or trailing text after last chapter and retry!", vbOKOnly
    Exit Sub
End If


If Err.Number <> 0 Then
    MsgBox "Err " & Err.Number & " occured in " & Err.Source & ": " & Err.Description
    Err.Clear
    Exit Sub
End If


Dim tgDoc As Document
Dim orDoc As Document

' solve documents puzzle
If InStr(1, Documents(1).Name, ".en") > 0 Then
    Set orDoc = Documents(1)
    Set tgDoc = Documents(2)
Else
    Set orDoc = Documents(2)
    Set tgDoc = Documents(1)
End If

'***************************************************************************************************************************************
'************************************************* FIRST SOURCE DOC PROCESSING *********************************************************
'***************************************************************************************************************************************


' First original doc
'orDoc.Activate

Dim nrTopTables As Integer
' count its top tables
nrTopTables = Count_TopTables(orDoc)

Call All_Bookmarks_Remove(orDoc)

' process target doc for sorting
Call Set_Top_Bookmarks(orDoc, False)
' both bkm categories are needed
Call Set_MainBookmarks(orDoc, nrTopTables)

Dim OriginalDocument_Order_SortingKeys() As String

'retrieve original doc chapter numbers order (to be followed on sorting target language doc
OriginalDocument_Order_SortingKeys = Get_MainIDs_Original_SortingOrder_forAllBookmarks(orDoc, "Fado_", nrTopTables)


'***************************************************************************************************************************************
'*********************************************** SWITCH TO TARGET DOC PROCESSING *******************************************************
'***************************************************************************************************************************************


' switch active doc for preprocessing
'tgDoc.Activate

Dim tgTopTablesCount As Integer
tgTopTablesCount = Count_TopTables(tgDoc)

If tgTopTablesCount <> nrTopTables Then
    MsgBox "Warning ! Top tables count in target document is different from source document one! (" & tgTopTablesCount & "/ " & nrTopTables & ")", vbOKOnly + vbCritical
    Exit Sub
End If


Call All_Bookmarks_Remove(tgDoc)

'
Call Set_Top_Bookmarks(tgDoc, False)

'
Call Set_MainBookmarks_accToSortingOrder(tgDoc, OriginalDocument_Order_SortingKeys, tgTopTablesCount)


' TO DO
Call Bkm_Sort_byName(tgDoc, tgTopTablesCount)


End Sub


Sub Set_Top_Bookmarks(TargetDocument As Document, Remove_MultiRow_Tables As Boolean)
' Set all topS### bookmarks, one jest before all chapters' top table.

Dim tt As Table
Dim topCount As Integer

Dim tempTopB As Bookmark
Dim tempRange As Range

For Each tt In TargetDocument.Tables
    If tt.Rows.Count = 1 Then
        If tt.Rows(1).Cells.Count = 1 Then
            If tt.Shading.BackgroundPatternColorIndex = 16 Then
                'tt.Rows(1).Cells(1).Range.Text Like "??top??" Or _
                tt.Rows(1).Cells(1).Range.Text Like "???nceput??" Or _
                tt.Rows(1).Cells(1).Range.Text Like "??nceput??" Then
                topCount = topCount + 1
                ' Our "top" tables !
                'tt.Range.Shading.BackgroundPatternColorIndex = wdYellow
                
                Set tempRange = tt.Range
                tempRange.Collapse (wdCollapseStart)
                tempRange.MoveEnd wdCharacter, -1
                TargetDocument.Bookmarks.Add "topS" & Format(topCount, "000"), tempRange
                
            End If
        Else
            ' To be converted to text and delete resulting pictures
            'tt.Range.Shading.BackgroundPatternColorIndex = wdPink
            'tt.ConvertToText (vbTab)
        End If
    Else
        If Remove_MultiRow_Tables Then
            ' To be converted to text and delete resulting pictures
            tt.Range.Shading.BackgroundPatternColorIndex = wdPink
            tt.ConvertToText (vbTab)
        End If
    End If
Next tt

MsgBox "Au fost setate " & topCount & " bookmarks la tabelele TOP!", vbOKOnly, "Numar capitole"

End Sub


Sub Set_MainBookmarks(TargetDocument As Document, HowManyTopTables As Integer)
' Set all main bookmarks, each containing one chapter, including the "top table" and finishing before
' the next top table's "before paragraph"

Dim tbk As Bookmark
Dim trng As Range
Dim hm As Integer

hm = HowManyTopTables

If TargetDocument.Bookmarks.Count > 0 Then
    
    For j = 1 To TargetDocument.Bookmarks.Count
        
        If Left$(TargetDocument.Bookmarks(j).Name, 4) = "topS" Then
            
            k = k + 1   ' topS bookmark index
            If k < hm Then
                Set trng = TargetDocument.Range(TargetDocument.Bookmarks("topS" & Format(k, "000")).Range.Start, _
                    TargetDocument.Bookmarks("topS" & Format(k + 1, "000")).Range.Start)
                ' "topS" & format(j+1, "000")).
                TargetDocument.Bookmarks.Add "Fado_" & Format(k, "000"), trng
                
            End If
        
        End If
        
    Next j
    
End If


MsgBox "Au fost setate " & k & " bookmarkuri <Fado_>"


End Sub

Sub Set_MainBookmarks_accToSortingOrder(TargetDocument As Document, OriginalDoc_SortingOrder() As String, HowManyTopTables As Integer)

' Set all main bookmarks, each containing one chapter, including the "top table" and finishing before
' the next top table's "before paragraph". Sorting order is the one corresponding to the original doc sort order, found in parameter OriginalDoc_SortingOrder

Dim tbk As Bookmark
Dim trng As Range
Dim hm As Integer

hm = HowManyTopTables


Dim ctChapter_range As Range
Dim ctChapter_tmpBkm As Bookmark                ' temporary bookmark of ct chapter range
Dim ctChapter_mainID As String
Dim ctChapter_mainID_origPosition As Integer


If TargetDocument.Bookmarks.Count > 0 Then
    
    For j = 1 To HowManyTopTables
        
        If Left$(TargetDocument.Bookmarks(j).Name, 4) = "topS" Then
            
            k = k + 1   ' topS bookmark index
            
            If k < hm Then
                
                Set ctChapter_range = TargetDocument.Range(TargetDocument.Bookmarks("topS" & Format(k, "000")).Range.Start, _
                    TargetDocument.Bookmarks("topS" & Format(k + 1, "000")).Range.Start)
                
                ' to remove later, temp bookmark only
                Set ctChapter_tmpBkm = TargetDocument.Bookmarks.Add("tmpBkm_" & k, ctChapter_range)
                
                ctChapter_mainID = Get_MainID_fromBookmark(TargetDocument, ctChapter_tmpBkm)
                
                ctChapter_mainID_origPosition = gscrolib.Get_Index_ofArray_Entry(OriginalDoc_SortingOrder, ctChapter_mainID)
                
                ' "topS" & format(j+1, "000")).
                TargetDocument.Bookmarks.Add "Fado_" & Format(ctChapter_mainID_origPosition, "000"), ctChapter_range
                
            End If
        
        End If
        
    Next j
    
End If


MsgBox "Au fost setate " & k & " bookmarkuri <Fado_>"



End Sub


Function Get_MainIDs_Original_SortingOrder_forAllBookmarks(TargetDocument As Document, BookmarksPrefix As String, NumberOfBookmarks As Integer) As String()

Dim tmpKeys() As String

' TO BE RAN supplying original doc number, create array of main chapter IDs in order of EN doc

' Similar to Create_Fado_SortingKeys_ForAllBookmarks, which is essential to sorting in linguistic order.
' This one is however used to create an array of strings representing the main chapter IDs for the FADO glossary
' This array represents the order the chapters need to be in to be sorted same as EN, needed therefore for post-alignment
' if document has been done without Trados !

' Possibly necessary also for preparation of a FADO glossary base, when we need to resort a lingustically sorted doc in EN order, for alignment in Euramis

Dim tfn_par As String

Dim ck_tbk As Bookmark

If TargetDocument.Bookmarks.Count > 0 Then
    
    For k = 1 To NumberOfBookmarks
        
        If TargetDocument.Bookmarks.Exists(BookmarksPrefix & Format(k, "000")) Then
            
            Set ck_tbk = TargetDocument.Bookmarks(BookmarksPrefix & Format(k, "000"))
            
            ReDim Preserve tmpKeys(k)
            
            tfn_par = Get_MainID_fromBookmark(TargetDocument, ck_tbk)
            
            If tfn_par <> "" Then
                tmpKeys(k) = tfn_par
            Else
                MsgBox "Getting main ID paragraph in " & ck_tbk.Name & " bookmark" & _
                    " failed, returning string was empty, please investigate !", vbOKOnly + vbCritical, "Error, no key for current bookmark!"
                Exit Function
            End If
            
        Else
            MsgBox "Please check bookmarks, bookmark " & BookmarksPrefix & Format(k, "000") & _
                " does not exist !", vbOKOnly + vbCritical, "Necessary bookmark not found"
            Exit Function
        End If
        
    Next k
    
Else
    MsgBox "No bookmarks found in Document, Exiting !", vbOKOnly + vbCritical, "Creating sorting keys failed!"
    Exit Function
End If

'If UBound(tmpKeys) > 0 Then
    'MsgBox "The tmpKeys array has now been filled with " & UBound(strKeys) & " main IDs sorting keys !"
'End If


Get_MainIDs_Original_SortingOrder_forAllBookmarks = tmpKeys

End Function


Sub Create_Fado_SortingKeys_ForAllBookmarks(BookmarksPrefix As String, NumberOfBookmarks As Integer)

Dim tfn_par As String
Dim ck_tbk As Bookmark

If ActiveDocument.Bookmarks.Count > 0 Then
    For k = 1 To NumberOfBookmarks
        If ActiveDocument.Bookmarks.Exists(BookmarksPrefix & Format(k, "000")) Then
            Set ck_tbk = ActiveDocument.Bookmarks(BookmarksPrefix & Format(k, "000"))
            
            ReDim Preserve strKeys(k)
            
            tfn_par = Get_First_NonEmpty_NonNumeric_Paragraph(ck_tbk)
            
            If tfn_par <> "" Then
                strKeys(k) = LCase(tfn_par)
            Else
                MsgBox "Getting first non empty paragraph in " & ck_tbk.Name & " bookmark" & _
                    " failed, returning string was empty, please investigate !", vbOKOnly + vbCritical, "Error, no key for current bookmark!"
                Exit Sub
            End If
            
        Else
            MsgBox "Please check bookmarks, bookmark " & BookmarksPrefix & Format(k, "000") & _
                " does not exist !", vbOKOnly + vbCritical, "Necessary bookmark not found"
            Exit Sub
        End If
    Next k
    
Else
    MsgBox "No bookmarks found in Document, Exiting !", vbOKOnly + vbCritical, "Creating sorting keys failed!"
    Exit Sub
End If

If UBound(strKeys) > 0 Then
    MsgBox "The strKeys array has now been filled with " & UBound(strKeys) & " sorting keys !"
End If

End Sub


Sub Bkm_Sort_byName(TargetDocument As Document, BookmarksNumber As Integer)


If TargetDocument.Bookmarks.Count = 0 Then
    MsgBox "Document " & TargetDocument.Name & " had NO bookmarks, Exiting..."
    Exit Sub
End If

Dim firstTopStartRange As Range

If TargetDocument.Bookmarks.Exists("topS001") Then
    Set firstTopStartRange = TargetDocument.Bookmarks("topS001").Range
Else
    MsgBox "Ooops!"
    Exit Sub
End If


Dim bkmsrc As Bookmark
Dim rngsrc As Range
Dim rngdst As Range


Set rngdst = firstTopStartRange.Duplicate

Load frmSortPages

For i = 1 To BookmarksNumber
    
    frmSortPages.lblMessage = "Sorting text block " & i & " on " & BookmarksNumber
    
    If TargetDocument.Bookmarks.Exists("Fado_" & Format(i, "000")) Then
        
        Set bkmsrc = TargetDocument.Bookmarks("Fado_" & Format(i, "000"))
        
        Set rngsrc = bkmsrc.Range.Duplicate
        rngsrc.Cut
        
        
        rngdst.Paste

        'bkmdst.Start = rngdst.End
        
        rngdst.Collapse (wdCollapseEnd)
        
    Else
    End If
    
    DoEvents
    ActiveDocument.UndoClear
    
Next i

Unload frmSortPages


Application.ScreenRefresh

If i = (BookmarksNumber + 1) Then
    MsgBox "Sorted " & i - 1 & " blocks of text", vbOKOnly + vbInformation, "Sorting successfully finished!"
End If

End Sub



Sub Bkm_Sort(BookmarksPrefix As String, Optional BookmarksNumber)
' User optional argument if doc contains other bookmarks as well, besides those we want sorted
' otherwise nrd variable is set to number of bookmarks in document

Dim bkmsrc As Bookmark, bkmdst As Bookmark
Dim rngsrc As Range, rngdst As Range
Dim i As Integer, j As Integer, nctd As Integer
Dim nrd As Integer

Dim mainTRange As Range
Set mainTRange = ActiveDocument.StoryRanges(wdMainTextStory)
' We can't sort using ActiveDocument.Bookmarks, since the index order is NOT reset after you move bookmarks around for the Document object,
' apparently intentionally. IT IS fot the range/ selection object !!!!!

System.Cursor = wdCursorWait

If Not IsMissing(BookmarksNumber) Then
    nrd = BookmarksNumber
Else
    nrd = mainTRange.Bookmarks.Count
End If

For i = 2 To nrd
    frmSortPages.lblMessage = "Sorting text block " & i & " on " & nrd
    
    'If ActiveDocument.Bookmarks.Exists(BookmarksPrefix & Format(i, "000")) Then
        'Set bkmsrc = ActiveDocument.Bookmarks(BookmarksPrefix & Format(i, "000"))
    'Else
        'MsgBox "Please check bookmarks, bookmark " & _
            'BookmarksPrefix & Format(i, "000") & " does not exist !"
        'Exit Sub
    'End If
    
    Set bkmsrc = mainTRange.Bookmarks(i)

    
    'bkmsrc.Select
    For j = 1 To i - 1
        'If ActiveDocument.Bookmarks.Exists(BookmarksPrefix & Format(j, "000")) Then
            'Set bkmdst = ActiveDocument.Bookmarks(BookmarksPrefix & Format(j, "000"))
        'Else
            'MsgBox "Please check bookmarks, bookmark " & _
                BookmarksPrefix & Format(j, "000") & " does not exist !"
            'Exit Sub
        'End If
        
        Set bkmdst = mainTRange.Bookmarks(j)
        
        If StrComp(strKeys(CInt(Right(bkmsrc.Name, 3))), strKeys(CInt(Right(bkmdst.Name, 3))), vbTextCompare) = -1 Then       ' We use three Digits for bookmarks numbers
            Set rngsrc = bkmsrc.Range.Duplicate
            rngsrc.Cut
            Set rngdst = bkmdst.Range.Duplicate
            rngdst.Collapse wdCollapseStart
            rngdst.Paste
            bkmdst.Start = rngdst.End
            rngdst.Collapse (wdCollapseEnd)
            'nctd = rngdst.MoveStartWhile(Chr(32), wdBackward)
            'rngdst.MoveStartUntil vbCr, wdBackward
            'rngdst.Text = ""
            Exit For
        End If
        
        DoEvents
        System.Cursor = wdCursorWait
    Next j
    ActiveDocument.UndoClear
Next i

Unload frmSortPages

' we remove final unnecessary pagebreak and empty cr's
' Call Clean_RngSrc_Ends
Application.ScreenRefresh

If i = (nrd + 1) Then
    MsgBox "Sorted " & i - 1 & " blocks of text", vbOKOnly + vbInformation, "Sorting successfully finished!"
End If

'Call clean_up_sort

Set bkmsrc = Nothing: Set bkmdst = Nothing
Set rngsrc = Nothing: Set rngdst = Nothing
i = 0: j = 0

End Sub


Sub All_OLEShapes_Count()

Dim tgDoc As Document
Dim sdoc As Document

Set sdoc = ActiveDocument
Set tgDoc = Documents.Add
sdoc.Activate

Dim ctIS As InlineShape

Dim counter As Integer

For Each ctIS In ActiveDocument.InlineShapes
    If ctIS.Type = wdInlineShapeEmbeddedOLEObject Then
        counter = counter + 1
        
        'ctIS.Range.Copy
        'tgdoc.StoryRanges(wdMainTextStory).InsertParagraphAfter
        'tgdoc.Paragraphs(tgdoc.Paragraphs.Count).Range.Paste
        
    End If
Next ctIS

MsgBox "Counted " & counter & " inline embedded OLEobjects", vbOKOnly + vbInformation, "Done"


End Sub

Sub AllOleShapes_Extract()

Dim tgDoc As Document
Dim sdoc As Document

Set sdoc = ActiveDocument
Set tgDoc = Documents.Add
sdoc.Activate

Dim ctIS As InlineShape

Dim counter As Integer

For Each ctIS In ActiveDocument.InlineShapes
    If ctIS.Type = wdInlineShapeEmbeddedOLEObject Then
        counter = counter + 1
        
        ctIS.Range.Copy
        tgDoc.StoryRanges(wdMainTextStory).InsertParagraphAfter
        
        tgDoc.Paragraphs(tgDoc.Paragraphs.Count).Range.Paste
        
    End If
Next ctIS

MsgBox "Extracted " & counter & " inline embedded OLEobjects", vbOKOnly + vbInformation, "Done"

End Sub

Sub AllOle_InlineShapes_ConvertToPicture()

'Dim tgdoc As Document
'Dim sdoc As Document
'
'Set sdoc = ActiveDocument
'Set tgdoc = Documents.Add
'sdoc.Activate

Dim ctIS As InlineShape

Dim counter As Integer

For Each ctIS In ActiveDocument.InlineShapes
    If ctIS.Type = wdInlineShapeEmbeddedOLEObject Then
        
        counter = counter + 1
        
        cits.Select
        ctIS.Type = wdInlineShapePicture
        
    End If
Next ctIS

MsgBox "Converted " & counter & " inline embedded OLEobjects to inline shapes!", vbOKOnly + vbInformation, "Done"



End Sub

Sub Highlight_All_TopTables_CvTxt_Rest()

Dim tt As Table


If ActiveDocument.Tables.Count > 0 Then
    
    Dim topTablesCount As Integer
    Dim mrTablesCount As Integer
    
    For Each tt In ActiveDocument.Tables
        
        If tt.Rows.Count = 1 Then
            
            If tt.Rows(1).Cells.Count = 1 Then
                If tt.Shading.BackgroundPatternColorIndex = 16 Then
                    topTablesCount = topTablesCount + 1
                    tt.Range.HighlightColorIndex = wdYellow
                End If
            End If
            
        Else    ' Multirow
            
            mrTablesCount = mrTablesCount + 1
            
            tt.Range.HighlightColorIndex = wdPink
            
            tt.ConvertToText
            
        End If
        
    Next tt
    
    MsgBox "Highlighted " & topTablesCount & " top tables, highlighted & converted txt " & mrTablesCount & " multiple rows tables"
    
Else
    MsgBox "Found NO tables in " & ActiveDocument.Name
End If


End Sub


Function Count_TopTables(TargetDocument As Document) As Integer


Dim tt As Table

Dim topCount As Integer


For Each tt In TargetDocument.Tables
    If tt.Rows.Count = 1 Then
        If tt.Rows(1).Cells.Count = 1 Then
            If tt.Shading.BackgroundPatternColorIndex = 16 Then
                'tt.Rows(1).Cells(1).Range.Text Like "??top??" Or _
                tt.Rows(1).Cells(1).Range.Text Like "???nceput??" Or _
                tt.Rows(1).Cells(1).Range.Text Like "??nceput??" Then
                topCount = topCount + 1
                ' Our "top" tables !
                'tt.Range.Shading.BackgroundPatternColorIndex = wdYellow
                
                Set tempRange = tt.Range
                tempRange.Collapse (wdCollapseStart)
                tempRange.MoveEnd wdCharacter, -1
                TargetDocument.Bookmarks.Add "topS" & Format(topCount, "000"), tempRange
            End If
        End If
    End If
Next tt

Count_TopTables = topCount

MsgBox topCount & " tabele TOP in " & TargetDocument.Name, vbOKOnly, "Numar capitole"


End Function

Sub Remove_AllTopTables()

Dim tt As Table

Dim topCount As Integer


For Each tt In ActiveDocument.Tables
    If tt.Rows.Count = 1 Then
        If tt.Rows(1).Cells.Count = 1 Then
            If tt.Shading.BackgroundPatternColorIndex = 16 Then
                'tt.Rows(1).Cells(1).Range.Text Like "??top??" Or _
                tt.Rows(1).Cells(1).Range.Text Like "???nceput??" Or _
                tt.Rows(1).Cells(1).Range.Text Like "??nceput??" Then
                topCount = topCount + 1
                ' Our "top" tables !
                'tt.Range.Shading.BackgroundPatternColorIndex = wdYellow
                
                tt.Delete
                
            End If
        End If
    End If
Next tt

MsgBox "Removed " & topCount & " tabele TOP in " & ActiveDocument.Name, vbOKOnly

End Sub


Sub Empty_NonInlineShape_HyperLinks_Delete()

Dim chyp As Hyperlink
Dim counter As Integer

If ActiveDocument.Hyperlinks.Count > 0 Then
    For Each chyp In ActiveDocument.Hyperlinks
        If chyp.TextToDisplay = "" Then
            'chyp.Range.Select
            
            If Selection.Type <> wdSelectionInlineShape Then
                chyp.Delete
                counter = counter + 1
            Else
                                
            End If
            
            'MsgBox "Found one Empty hyperlink !"
        End If
    Next chyp
End If


If counter > 0 Then
    MsgBox "Au fost gasite si eliminate " & counter & " hyperlinkuri goale !"
End If

End Sub



Sub AllTables_FitContents()

Dim ctb As Table

For Each ctb In ActiveDocument.Tables
    ctb.AllowAutoFit = True
Next ctb

End Sub

Sub AllReferences_InMainText_ToRefStyle()

Dim ctfn As Footnote

If ActiveDocument.Footnotes.Count > 0 Then
    For Each ctfn In ActiveDocument.Footnotes
        ctfn.Reference.Style = "Footnote Reference"
    Next ctfn
End If

End Sub

Sub CoverNote_Remove()

If ActiveDocument.Sections.Count > 0 Then
    If ActiveDocument.Sections(1).Range.Information(wdActiveEndPageNumber) = 2 Then
        ActiveDocument.Sections(1).Range.Delete
    End If
End If

End Sub

Sub Empty_OneRow_Tables_Remove()

Dim howManyET As Integer
Dim ttab As Table

If ActiveDocument.Tables.Count > 0 Then
    For Each ttab In ActiveDocument.Tables
        If ttab.Rows.Count = 1 Then
            If IsTable_Empty(ttab) Then
                howManyET = howManyET + 1
                ttab.Delete
            End If
        End If
    Next ttab
End If

If howManyET > 0 Then
    MsgBox "Found and removed " & howManyET & " one row empty tables !"
Else
    MsgBox "Found NO empty multi-row tables !"
End If

End Sub

Sub Empty_MultiRowTables_Count()

Dim howManyET As Integer
Dim ttab As Table

If ActiveDocument.Tables.Count > 0 Then
    For Each ttab In ActiveDocument.Tables
        If ttab.Rows.Count > 1 Then
            If IsTable_Empty(ttab) Then
                ttab.Range.Select
                howManyET = howManyET + 1
                'ttab.Delete
            End If
        End If
    Next ttab
End If

If howManyET > 0 Then
    MsgBox "Found " & howManyET & " one row empty tables !"
Else
    MsgBox "Found NO empty multi-row tables !"
End If


End Sub

Sub Empty_MultiRowTables_Remove()

Dim howManyET As Integer
Dim ttab As Table

If ActiveDocument.Tables.Count > 0 Then
    For Each ttab In ActiveDocument.Tables
        If ttab.Rows.Count > 1 Then
            If IsTable_Empty(ttab) Then
                ttab.Range.Select
                howManyET = howManyET + 1
                ttab.Delete
            End If
        End If
    Next ttab
End If

If howManyET > 0 Then
    MsgBox "Found and removed " & howManyET & " one row empty tables !"
Else
    MsgBox "Found NO empty multi-row tables !"
End If


End Sub

Function IsTable_Empty(TargetTable As Table) As Boolean

Dim ttab As Table
Set ttab = TargetTable

Dim realtext As String

realtext = ttab.Range.Text
realtext = Replace(Replace(Replace(realtext, vbCr, ""), vbLf, ""), vbCrLf, "")
realtext = Replace(Replace(Replace(realtext, vbcrcr, ""), " ", ""), Chr(160), "")
realtext = Replace(Replace(realtext, Chr(13), ""), Chr(7), "")

If realtext <> "" Then
    IsTable_Empty = False
Else
    IsTable_Empty = True
End If
        
End Function

Sub AllVBCr_Pairs_ToVbCr()
' Clean empty paragraphs preceded by other "space symbols"
' such as Tab+Enter, Space+Enter, HardSpace+Enter

Dim ls As String
ls = gscrolib.GetListSeparatorFromRegistry

If ls = "" Then MsgBox "List separator not extracted from registry!", vbOKOnly, "Please debug !": Exit Sub

With ActiveDocument.StoryRanges(wdMainTextStory).Find
    .ClearAllFuzzyOptions
    .ClearFormatting
    .Format = False
    .Forward = True
    .Wrap = wdFindStop
        
    .Execute FindText:="([ ^t^s]{1" & ls & "})(^13)", MatchWildcards:=True, Forward:=True, Wrap:=wdFindStop, Format:=False, _
        ReplaceWith:="^p", Replace:=wdReplaceAll
End With


End Sub

Sub AllMultiple_VBCrs_To_Single()

Dim ls As String
ls = gscrolib.GetListSeparatorFromRegistry

If ls = "" Then MsgBox "List separator not extracted from registry!", vbOKOnly, "Please debug !": Exit Sub

With ActiveDocument.StoryRanges(wdMainTextStory).Find
    .ClearAllFuzzyOptions
    .ClearFormatting
    .Format = False
    .Forward = True
    .Wrap = wdFindStop
        
    .Execute FindText:="(^13{2" & ls & "})", MatchWildcards:=True, Forward:=True, Wrap:=wdFindStop, Format:=False, _
        ReplaceWith:="^p", Replace:=wdReplaceAll
End With


End Sub



Function Get_First_NonEmpty_NonNumeric_Paragraph(TargetBookmark As Bookmark) As String

Dim tpar As Paragraph
Dim realtext As String

If ActiveDocument.Bookmarks.Count > 0 Then
    If ActiveDocument.Bookmarks.Exists(TargetBookmark) Then
        For Each tpar In TargetBookmark.Range.Paragraphs
                       
            ' Trying to remove all expectable non-printing characters, to retrieve text only as it were
            realtext = tpar.Range.Text
            realtext = Replace(Replace(Replace(realtext, vbCr, ""), vbLf, ""), vbCrLf, "")
            realtext = Replace(Replace(Replace(realtext, Chr(12), ""), Chr(160), " "), "  ", " ")
            realtext = Replace(Replace(Replace(realtext, Chr(10), ""), Chr(13), ""), Chr(7), "")
            realtext = Replace(Replace(realtext, Chr(1), ""), Chr(13), "")
            
            If realtext <> "" And realtext <> Chr(1) & "inceput" And _
               (Not realtext Like "(#*#)" And Not realtext Like " ###*" And _
               Not realtext Like "?top" And Not realtext Like "?(#*#)") Then
                Get_First_NonEmpty_NonNumeric_Paragraph = realtext
                Exit For
            Else
                
            End If
            
        Next tpar
    Else
        Get_First_NonEmpty_NonNumeric_Paragraph = ""
        Exit Function
    End If
Else
    Get_First_NonEmpty_NonNumeric_Paragraph = ""
    Exit Function
End If

End Function

Sub ExtractAll_NonEmpty_NonNumeric_FirstParagraphs_ToNewDoc()   ' Effectively, extract Index (contents table) texts (chapter's titles)

Dim tbk As Bookmark
Dim tdoc As Document
Dim ctdoc As Document
Dim counter As Integer

Dim trange As Range     ' Target Range
Set trange = ActiveDocument.StoryRanges(wdMainTextStory)

Dim k As Integer

Set ctdoc = ActiveDocument

For j = 1 To trange.Bookmarks.Count
    Set tbk = trange.Bookmarks(j)
    
    If Left$(tbk.Name, 4) = "Fado" Then
        k = k + 1   ' Fado bookmarks index
        If k = 1 Then
            Set tdoc = Documents.Add
            ctdoc.Activate
            
            tdoc.Content.InsertAfter (Get_First_NonEmpty_NonNumeric_Paragraph(tbk)) & vbCr
            counter = counter + 1
        Else
            tdoc.Content.InsertAfter (Get_First_NonEmpty_NonNumeric_Paragraph(tbk)) & vbCr
            counter = counter + 1
        End If
    End If
Next j

If counter > 0 Then
    MsgBox "S-au gasit si colectat " & counter & " prime paragrafe ne-goale"
End If

End Sub

Function Get_MainID_fromBookmark(TargetDocument As Document, TargetBookmark As Bookmark) As String

Dim tpar As Paragraph
Dim mainID As String
Dim tres As String      ' temporary result

If TargetDocument.Bookmarks.Count > 0 Then
    If TargetDocument.Bookmarks.Exists(TargetBookmark) Then
        
        For Each tpar In TargetBookmark.Range.Paragraphs
                       
            ' Trying to remove all expectable non-printing characters, to retrieve text only as it were
            mainID = tpar.Range.Text
            mainID = Replace(mainID, vbCr, "")
            'mainID = Replace(Replace(Replace(mainID, vbCr, ""), vbLf, ""), vbCrLf, "")
            'mainID = Replace(Replace(Replace(mainID, Chr(12), ""), Chr(160), " "), "  ", " ")
            'mainID = Replace(Replace(Replace(mainID, Chr(10), ""), Chr(13), ""), Chr(7), "")
            
            If mainID Like Chr(160) & "###*" Then
                tres = mainID
                
                If InStr(1, tres, ",") > 0 Then
                    Get_MainID_fromBookmark = Split(tres, ",")(0)   ' Chapter has actually more than one main ID, separated by comma
                Else
                    Get_MainID_fromBookmark = tres
                End If
                
                Exit For
                
            Else
            End If
            
        Next tpar
            
    Else
        Get_MainID_fromBookmark = ""
    End If
Else
    Get_MainID_fromBookmark = ""
End If

End Function

Sub GetAll_MainIDs_ToNewDocument_Name()      ' GET BOOKMARKS IN NAME ORDER, POSITION SET ASIDE (FOR SITUATION WHEN THEY ARE ARRANGED IN NAME-ORDER)
' More complete verions, extract chapter name, numeric ID & bookmark name

Dim tbk As Bookmark
Dim tdoc As Document
Dim cdoc As Document
Dim counter As Integer

Dim k As Integer

Dim trange As Range
Set trange = ActiveDocument.StoryRanges(wdMainTextStory)

Set cdoc = ActiveDocument

If trange.Bookmarks.Count > 0 Then
    For j = 1 To trange.Bookmarks.Count
        
        Set tbk = trange.Bookmarks(j)
        
        If Left$(tbk.Name, 5) = "Fado_" Then
            k = k + 1       ' good bookmarks index
            If k = 1 Then
                Set tdoc = Documents.Add
                cdoc.Activate
                
                tdoc.Content.InsertAfter ("Bkm " & tbk.Name & "   =   " & _
                    Get_First_NonEmpty_NonNumeric_Paragraph(tbk) & "   " & vbTab & "-   ID " & Get_MainID_fromBookmark(tbk) & vbCr)
                counter = counter + 1
            Else
                tdoc.Content.InsertAfter ("Bkm " & tbk.Name & "   =   " & _
                    Get_First_NonEmpty_NonNumeric_Paragraph(tbk) & "   " & vbTab & "-   ID " & Get_MainID_fromBookmark(tbk) & vbCr)
                counter = counter + 1
            End If
        End If
        
    Next j
End If

If counter > 0 Then
    MsgBox "S-au gasit si extras " & counter & " main IDs"
End If

End Sub

Sub Extract_AndList_AllHyperlinks()
' Put into new document all fields of type hyperlink, list their sub-address

Dim tbk As Bookmark
Dim counter As Integer, bcounter As Integer, idx As Integer

Dim sdoc As Document, tdoc As Document
Set sdoc = ActiveDocument

Dim thy As Hyperlink
Dim ttb As Table

Dim tsubs As String

If ActiveDocument.Bookmarks.Count > 0 Then
    For Each tbk In ActiveDocument.StoryRanges(wdMainTextStory).Bookmarks   ' accessing by range provides for ordered access
        
        If Left$(tbk.Name, 5) = "Fado_" Then    ' Process only the good hyperlinks
        
            If counter = 0 Then
                idx = idx + 1
                Set tdoc = Documents.Add: tdoc.Tables.Add tdoc.Paragraphs(1).Range, 1, 3: Set ttb = tdoc.Tables(1): sdoc.Activate
                
                ttb.Rows(1).Cells(1).Range.Text = "Main ID": ttb.Rows(1).Cells(2).Range.Text = "Chapter title": ttb.Rows(1).Cells(3).Range.Text = "Hyperlinks and targets"
                
                ttb.Rows.Add
                ttb.Rows(2).Cells(1).Range.Text = Get_MainID_fromBookmark(tbk)
                ttb.Rows(2).Cells(2).Range.Text = Get_First_NonEmpty_NonNumeric_Paragraph(tbk)
                            
                If tbk.Range.Hyperlinks.Count > 0 Then
                        
                    counter = counter + 1
                    
                    For Each thy In tbk.Range.Hyperlinks
                        tsubs = tsubs & thy.SubAddress & vbTab
                    Next thy
                    
                    ttb.Rows(2).Cells(3).Range.Text = "Hyperlinks: " & tbk.Range.Hyperlinks.Count & vbCr & tsubs
                    tsubs = ""
                                                    
                Else
                    bcounter = bcounter + 1
                    ttb.Rows(2).Cells(3).Range.Text = "Hyperlinks: " & tbk.Range.Hyperlinks.Count
                End If
                
            Else
                idx = idx + 1
                
                ttb.Rows.Add
                
                ttb.Rows(ttb.Rows.Count).Cells(1).Range.Text = Get_MainID_fromBookmark(tbk)
                ttb.Rows(ttb.Rows.Count).Cells(2).Range.Text = Get_First_NonEmpty_NonNumeric_Paragraph(tbk)
                            
                If tbk.Range.Hyperlinks.Count > 0 Then
                        
                    counter = counter + 1
                    
                    For Each thy In tbk.Range.Hyperlinks
                        tsubs = tsubs & thy.SubAddress & vbTab
                    Next thy
                    
                    ttb.Rows(ttb.Rows.Count).Cells(3).Range.Text = "Hyperlinks: " & tbk.Range.Hyperlinks.Count & vbCr & tsubs
                    tsubs = ""
                                                    
                Else
                    bcounter = bcounter + 1
                    ttb.Rows(ttb.Rows.Count).Cells(3).Range.Text = "Hyperlinks: " & tbk.Range.Hyperlinks.Count
                End If
            
            End If
        
        End If  ' If tbk.name este "Fado_###"
        
    Next tbk

tdoc.SpellingChecked = True    ' Get rid of pesky red underlining

End If


If counter > 0 Then
    If bcounter = 0 Then
        MsgBox "S-au procesat " & counter & " bookmarkuri cu cel putin 1 hiperlinkuri"
    Else
        MsgBox "Au fost gasite " & bcounter & " bookmarkuri fara hiperlinkuri"
    End If
    
End If

End Sub

Sub Build_Index_AtCursorPosition()
' Builds the "manual" contents table of the document at cursor position,
' by extracting chapters titles from bookmarks (Fado_###, # being digits)
' Make sure such bookmarks exist and contain a chapter each before running this one

Dim tbk As Bookmark
Dim counter As Integer

Dim trange As Range
Set trange = ActiveDocument.StoryRanges(wdMainTextStory)

If trange.Bookmarks.Count > 0 Then
    For j = 1 To trange.Bookmarks.Count
        
        Set tbk = trange.Bookmarks(j)
        
        If Left$(tbk.Name, 5) = "Fado_" Then
            
            Selection.InsertAfter Get_First_NonEmpty_NonNumeric_Paragraph(tbk)     ' To create hyperlink, insert text first
            ActiveDocument.Hyperlinks.Add Selection.Range, , "_" & _
                Replace(Get_MainID_fromBookmark(tbk), Chr(160), ""), , Selection.Text
            Selection.InsertAfter vbCr: Selection.Collapse (wdCollapseEnd)
            
            'selection.InsertAfter (Get_First_NonEmpty_NonNumeric_Paragraph(tbk) & Get_MainID_fromBookmark(tbk) & vbCr)
            counter = counter + 1
        
        End If
        
    Next j
End If

If counter > 0 Then
    MsgBox "S-au inserat " & counter & " intrari in Index !"
End If


End Sub

Sub All_MainIDs_ToHeading1_Style()

Dim tbk As Bookmark
Dim counter As Integer, bcounter As Integer

Dim trange As Range
Set trange = ActiveDocument.StoryRanges(wdMainTextStory)

Dim lookfor As String
Dim ttr As Range

If trange.Bookmarks.Count > 0 Then
    For j = 1 To trange.Bookmarks.Count
        
        Set tbk = trange.Bookmarks(j)
        
        lookfor = Get_MainID_fromBookmark(tbk)
        
                
        If lookfor <> "" Then
            
            Set ttr = tbk.Range
            
            If ttr.Find.Execute(FindText:=lookfor, ReplaceWith:="", Replace:=wdReplaceNone, Forward:=True, Wrap:=wdFindStop) Then
                counter = counter + 1
                
                ttr.Select: Selection.Expand (wdParagraph): Selection.ClearFormatting
                ttr.Paragraphs(1).Style = "Heading 1": ttr.ParagraphFormat.Alignment = wdAlignParagraphRight
                
            Else
                bcounter = bcounter + 1
            End If
                      
        End If
        
    Next j
End If

If counter > 0 Then
    If bcounter = 0 Then
        MsgBox "Am identificat si stiluit " & counter & " main IDs"
    Else
        MsgBox "Am identificat si stiluit " & counter & " main IDs" & vbcrcr & _
            bcounter & " bookmarkuri au rezultat negativ !"
    End If
End If


End Sub

Sub CurrentSelection_CrossRef_Add()

Dim tg As Range
Set tg = Selection.Range

If tg.Hyperlinks.Count > 0 Then
    
    tg.Hyperlinks(1).Delete
    
    tg.Hyperlinks.Add Address:="", Anchor:=tg, SubAddress:="_000", TextToDisplay:=""
    

    'ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
        SubAddress:="topS015", ScreenTip:="", TextToDisplay:="O"
    

End If



End Sub

Sub Replace_WrongSubAdress_to000_inAllHyperLinks(CurrentWrongAdress As String)
' Change all top links from tables to correct target, "_000"
' For some reason, mine all point to "_186"

Dim hp As Hyperlink
Dim counter As Integer

Dim trng As Range
Dim cta As String
cta = CurrentWrongAdress

If ActiveDocument.Hyperlinks.Count > 0 Then
    
    For Each hp In ActiveDocument.Hyperlinks
        
        If hp.SubAddress = cta Then
            
            Set trng = hp.Range
            hp.Delete
            trng.Hyperlinks.Add Address:="", Anchor:=trng, SubAddress:="_000", TextToDisplay:=""
            counter = counter + 1
        
        End If
    Next hp
End If

MsgBox "Changed " & counter & " hyperlinks to point to _000 bookmark target!", vbOKOnly, "Done!"

End Sub

Sub Remove_1_from_AllHyperlinks()
' In current document, no "_###_1" bookmarks have been used/ set.
' Remove the trailing "_1" from all hypers, in consequence.

Dim hp As Hyperlink
Dim counter As Integer

Dim trng As Range
Dim wadd As String      ' wrong address
wadd = "_1"

Dim cadd As String      ' correct address (same, without "_1")
    
If ActiveDocument.Hyperlinks.Count > 0 Then
    
    For Each hp In ActiveDocument.Hyperlinks
        
        If Right$(hp.SubAddress, 2) = wadd Then
            
            cadd = Replace(hp.SubAddress, wadd, "")     ' we remove "_1" tail by replacing it to nothing
            
            Set trng = hp.Range
            hp.Delete
            
            trng.Hyperlinks.Add Address:="", Anchor:=trng, SubAddress:=cadd, TextToDisplay:=""
            counter = counter + 1
        
        End If
    Next hp
Else
    MsgBox "Document does NOT contain ANY hyperlinks!", vbOKOnly, "NO"
    Exit Sub
End If

MsgBox "Removed trailing " & """ & _1 & """ & " from " & counter & " hyperlinks !", vbOKOnly, "Done!"



End Sub

Sub All_Bookmarks_WithPrefix_Remove(Prefix As String)

Dim rcount As Integer
Dim tbk As Bookmark

If ActiveDocument.Bookmarks.Count > 0 Then
    For Each tbk In ActiveDocument.Bookmarks
        If Left(tbk.Name, 4) = Prefix Then
            tbk.Delete
        End If
    Next tbk
End If

End Sub

Sub All_Bookmarks_Remove(TargetDocument As Document)

Dim bk As Bookmark

For Each bk In TargetDocument.Bookmarks
    bk.Delete
Next bk

End Sub

Sub RemoveAll_MultiRows_Tables()

Dim tt As Table

For Each tt In ActiveDocument.Tables
    If tt.Rows.Count > 1 Then
        tt.Delete
    End If
Next tt

End Sub

Sub Count_MultipleRows_Tables()


Dim tt As Table

Dim mrTables


For Each tt In ActiveDocument.Tables
    If tt.Rows.Count > 1 Then
        mrTables = mrTables + 1
    End If
Next tt

MsgBox "Found " & mrTables & " multiple rows tables in " & ActiveDocument.Name


End Sub


Sub AllHeadings1_IDS_Highlight()

Dim pp As Paragraph

For Each pp In ActiveDocument.Paragraphs
    If pp.Style = "Heading 1" Then
        If pp.Range.Text Like "*###*" Then
            pp.Range.HighlightColorIndex = wdDarkBlue
            pp.Range.Font.Color = wdColorWhite
        End If
    End If
Next pp

End Sub

Sub AllHeadings1_IDS_GreyArial14Bold()

Dim pp As Paragraph

For Each pp In ActiveDocument.Paragraphs
    If pp.Style = "Heading 1" Then
        If pp.Range.Text Like "?###*" Then
            pp.Range.Font.ColorIndex = wdGray50
            pp.Range.Font.Size = 14
            pp.Range.Font.Name = "Arial"
            pp.Range.Font.Bold = True
        End If
    End If
Next pp

End Sub

Sub All_Heading2_IDS_GreyArial14NoBold()

Dim pp As Paragraph

For Each pp In ActiveDocument.Paragraphs
    If pp.Style = "Heading 2" Then
        If pp.Range.Text Like "?###*" Then
            pp.Range.Font.ColorIndex = wdGray50
            pp.Range.Font.Size = 14
            pp.Range.Font.Name = "Arial"
            pp.Range.Font.Bold = False
        End If
    End If
Next pp


End Sub

Sub AllHeading1_IDS_Highlight() ' Bright green

Dim pp As Paragraph

For Each pp In ActiveDocument.Paragraphs
    If pp.Style = "Heading 1" Then
        If pp.Range.Text Like "?###*" Then
            pp.Range.HighlightColorIndex = wdBrightGreen
        End If
    End If
Next pp

End Sub

Sub AllHeading1_IDS_MakeBookmark_Before()
' Create bookmark called "_###" (where # is a digit) for all Main IDs with Heading1 Style
' Please make sure all main IDs (those pertaining to one chapter only, none of those secondarys)
' have Heading 1 style before running this one.
' CAREFUL for Heading 2 style IDs (secondary IDs) which could also be potential targets for hyperlinks.
' DEVELOP some way to detect those, COMPULSORY !

Dim pp As Paragraph

For Each pp In ActiveDocument.Paragraphs
    If pp.Style = "Heading 1" Then
        If pp.Range.Text Like "?###*" Then
            ActiveDocument.Bookmarks.Add "_" & Replace(Replace(pp.Range.Text, Chr(160), ""), vbCr, ""), ActiveDocument.Range(pp.Range.Start, pp.Range.Start)
        End If
    End If
Next pp


End Sub

Sub AllHeading2_IDS_Highlight() ' Turquoise

Dim pp As Paragraph

For Each pp In ActiveDocument.Paragraphs
    If pp.Style = "Heading 2" Then
        If pp.Range.Text Like "?###*" Then
            pp.Range.HighlightColorIndex = wdTurquoise
        End If
    End If
Next pp


End Sub

Sub Extract_AllHeadings1_Titles_ToNewDoc()
' NOT Completed. Use "Sub GetAll_MainIDs_ToNewDocument_Name"

Dim pp As Paragraph
Dim orDoc As Document

Set orDoc = ActiveDocument

Documents.Add


For Each pp In orDoc.Paragraphs
    If pp.Style = "Heading 1" Then
        'If Not pp.Range.Text Like "*###*" Then
            'pp.Range.HighlightColorIndex = wdDarkRed
            'pp.Range.Font.Color = wdColorWhite
            Selection.InsertAfter (pp.Range.Text)
        'End If
    End If
Next pp

End Sub

Sub Contents_Check_AndPutTogether()

Dim orDoc As Document
Dim rodoc As Document

Dim d As Document

For Each d In Documents
    If Left$(Split(d.Name, ".")(1), 2) = "en" Then
        Set orDoc = d
    ElseIf Left$(Split(d.Name, ".")(1), 2) = "ro" Then
        Set rodoc = d
    End If
Next d

Dim rotb As Table
Dim ortb As Table

Set rotb = rodoc.Tables(1)
Set ortb = orDoc.Tables(1)
    
' check if all numbers are there and nothing is extra
For i = 2 To ortb.Rows.Count
    If Val(ortb.Rows(i).Cells(1).Range.Text) = _
        Val(rotb.Rows(i).Cells(1).Range.Text) Then
                    
        ortb.Rows(i).Cells(3).Range.Text = _
            Left$(rotb.Rows(i).Cells(2).Range.Text, rotb.Rows(i).Cells(2).Range.Characters.Count - 1)
                    
    End If
    'Debug.Print Val(ortb.Rows(i).Cells(1).Range.Text) & vbTab & Val(rotb.Rows(i).Cells(1).Range.Text)
Next i

End Sub

Sub Swap_IDS_With_Titles()
' NOT Completed, not sure what its purpose was!

Dim pp As Paragraph

For Each pp In ActiveDocument.Paragraphs
    If pp.Style = "Heading 1" Then
        
    End If
    
Next pp

End Sub

Sub AllHeadings1_Highlight()        ' TURQUOISE

Dim p As Paragraph

For Each p In ActiveDocument.Paragraphs
    If p.Style = "Heading 1" Then
        p.Range.HighlightColorIndex = wdTurquoise   ' Heading 1 's
    End If
Next p

End Sub

Sub AllHeadings2_Highlight()        ' BRIGHT GREEN

Dim p As Paragraph

For Each p In ActiveDocument.Paragraphs
    If p.Style = "Heading 2" Then
        p.Range.HighlightColorIndex = wdBrightGreen   ' Heading 2 's
    End If
Next p

End Sub

Sub AllHeadings3_Highlight()       ' PINK

Dim p As Paragraph

For Each p In ActiveDocument.Paragraphs
    If p.Style = "Heading 3" Then
        p.Range.HighlightColorIndex = wdPink   ' Heading 3 's
    End If
Next p

End Sub

Sub AllParagraphs_WithBullet_ToHeading3() ' DO NOT RUN !! (Do manually)

Dim pp As Paragraph

For Each pp In ActiveDocument.Paragraphs
    If pp.Range.Characters(1).Font.Name = "Symbol" Then
        pp.Style = "Heading 3"
    End If
Next pp


End Sub

Sub All_Heading2_ToHeading1_Style_Change()

Dim p As Paragraph

For Each p In ActiveDocument.Paragraphs
    If p.Style = "Heading 2" Then
        p.Style = "Heading 1"
    End If
Next p

End Sub

Sub AllHeadings1_MarkIf_BeforeLessThen4Enters()

Dim p As Paragraph

For Each p In ActiveDocument
    If p.Style = "Heading 1" Then
        
    End If
Next p

End Sub

Sub AllTablesDelete()

Dim tt As Table

For Each tt In ActiveDocument.Tables
    tt.Delete
Next tt

End Sub

Sub AllTables_Remove()

Dim tbl As Table, tblcount As Integer

If Not Documents.Count > 0 Then   ' if document is opened
    If ActiveDocument.Tables.Count > 0 Then
        For Each tbl In ActiveDocument.Tables
            tblcount = tblcount + 1
            tbl.Delete
        Next tbl
        StatusBar = "All " & tblcount & " tables were removed!"
    End If
Else
    StatusBar = "Not working without opened document!"
End If

End Sub

Sub AllInlineImagesDelete()

Dim ishape As InlineShape

If Documents.Count > 0 Then
    If ActiveDocument.InlineShapes.Count > 0 Then
        For Each ishape In ActiveDocument.InlineShapes
            ishape.Range.Text = ""
        Next ishape
    Else
        StatusBar = "No inline shapes in current document!"
    End If
Else
    StatusBar = "Not working without opened document!"
End If

End Sub

Sub AllShapesDelete()

Dim shp As Shape

If Documents.Count > 0 Then
    If ActiveDocument.Shapes.Count > 0 Then
        For Each shp In ActiveDocument.Shapes
            shp.Delete
        Next ishape
    Else
        StatusBar = "No inline shapes in current document!"
    End If
Else
    StatusBar = "Not working without opened document!"
End If

Set shp = Nothing
End Sub

Sub AllPicturesDelete()

If Documents.Count > 0 Then
    If ActiveDocument.InlineShapes.Count > 0 Then
        Call AllInlineImagesDelete
    End If
    
    If ActiveDocument.Shapes.Count > 0 Then
        Call AllShapesDelete
    End If
End If

End Sub

Sub AllPageBreaks_ToEnters()

With ActiveDocument.StoryRanges(wdMainTextStory).Find
    .ClearFormatting
    .Execute "^12", , , True, , , True, wdFindStop, False, "^13", wdReplaceAll
End With

End Sub

Sub All_Heading1_OneEmptyParagraph_Before()
' Add an empty paragraph before each paragraph with Heading 1 style (to be later changed into "with any style"

Dim pp As Paragraph

For Each pp In ActiveDocument.Paragraphs
    If pp.Style = "Heading 1" Then
        pp.Range.InsertBefore (vbCr)
    End If
Next pp

End Sub

Sub AllHeadings1_BetweenENDoubleQuotes()

Dim pp As Paragraph
Dim pr As Range

For Each pp In ActiveDocument.Paragraphs
    If pp.Style = "Heading 1" Then
        Set pr = pp.Range: pr.MoveEndWhile vbCr, wdBackward
        pr.Text = ChrW(34) & pr.Text & ChrW(34)
    End If
Next pp

End Sub

Sub AllEmptyParagraphs_ToNormalStyle()

Dim pp As Paragraph

For Each pp In ActiveDocument.Paragraphs
    If pp.Style = "Heading 1" Then
        If pp.Range.Text = vbCr Then
            pp.Style = "Normal"
        End If
    End If
Next pp

End Sub

Sub AllHeading1_KeepWithNext()

Dim pp As Paragraph

For Each pp In ActiveDocument.Paragraphs
    If pp.Style = "Heading 1" Then
        pp.Format.KeepWithNext = True
    End If
Next pp

End Sub

Sub Show_DoubleQuotesIn_Headings1()
' Highlight PINK paragraphs styled Heading 1 which contain " character

Dim pp As Paragraph

For Each pp In ActiveDocument.Paragraphs
    If pp.Style = "Heading 1" Then
        If InStr(1, pp.Range.Text, ChrW(34)) > 0 Then
            pp.Range.HighlightColorIndex = wdPink
        End If
    End If
Next pp

End Sub

Sub AllHeadings1_NoQuotes()

Dim pp As Paragraph

For Each pp In ActiveDocument.Paragraphs
    If pp.Style = "Heading 1" Then
        pp.Range.Text = Replace(pp.Range.Text, ChrW(34), "")
        pp.Style = "Heading 1"
    End If
Next pp

End Sub

Sub Temporarily_DoubleQuotes_ToFrenchQuotes_InNonHeadings1()

Dim pp As Paragraph

For Each pp In ActiveDocument.Paragraphs
    If pp.Style <> "Heading 1" Then
        If InStr(1, pp.Range.Text, ChrW(34)) > 0 Then
            
        End If
    End If
Next pp

End Sub

Sub all20Sized_ToHeadings1()

Dim pp As Paragraph

For Each pp In ActiveDocument.Paragraphs
   If pp.Range.Words(1).Font.Size = 20 Then
    pp.Style = "Heading 1"
   End If
Next pp

End Sub

Sub countHeadings1()

Dim pp As Paragraph
Dim h1counter As Integer

For Each pp In ActiveDocument.Paragraphs
    If pp.Style = "Heading 1" Then
        h1counter = h1counter + 1
    End If
Next pp

Debug.Print ActiveDocument.Name & " has " & h1counter & " headings 1"

End Sub

'
' **********************************************************************************************************
' ***********   Routines for sorting blocks of text, made for ST 6781/12 REV1 & ST 9916/08 *****************
' **********************************************************************************************************
'
