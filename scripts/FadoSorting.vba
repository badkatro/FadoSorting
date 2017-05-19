Public strKeys() As String      ' string sorting keys array



'
' **********************************************************************************************************
' ***********   Routines for sorting blocks of text, made for ST 6781/12 REV1 & ST 9916/08 *****************
' ***********   That means for                  "Glosarul FADO"                             ****************
' **********************************************************************************************************
'

Sub Master_Process_Doc_For_Alignment(TargetDocument As Document)
' Also does sorting of "chapters" (identified by a header with a one row table containing word "top" and an arrow)

'Call CoverNote_Remove
' Bookmarks will be used for sorting, we need document to be clean for that
Call All_Bookmarks_Remove(TargetDocument)
' All little blue arrows are out, as well as most pictures (the ones from tables at least)
Call AllInlineImagesDelete(TargetDocument)
' All identifying "top" tables are highlighted (shaded in fact), rest of tables
' are converted to text
Call Highlight_All_TopTables_CvTxt_Rest(TargetDocument)
' Clean empty enters
Call AllVBCr_Pairs_ToVbCr(TargetDocument)
Call AllMultiple_VBCrs_To_Single(TargetDocument)
Call AllPageBreaks_ToEnters(TargetDocument)
Call RemoveAll_PagebreakBefores(TargetDocument)

Call Empty_MultiRowTables_Remove


End Sub


Sub Master_Process_CtDoc_For_Alignment()
' Also does sorting of "chapters" (identified by a header with a one row table containing word "top" and an arrow)

'Call CoverNote_Remove
' Bookmarks will be used for sorting, we need document to be clean for that
Call All_Bookmarks_Remove(ActiveDocument)

' All little blue arrows are out, as well as most pictures (the ones from tables at least)
Call AllInlineImagesDelete(ActiveDocument)

' All identifying "top" tables are highlighted (shaded in fact), rest of tables
' are converted to text
Call Highlight_All_TopTables_CvTxt_Rest(ActiveDocument)

' Clean empty enters
Call AllVBCr_Pairs_ToVbCr(ActiveDocument)
Call AllMultiple_VBCrs_To_Single(ActiveDocument)
Call AllPageBreaks_ToEnters(ActiveDocument)
Call RemoveAll_PagebreakBefores(ActiveDocument)

Call Empty_MultiRowTables_Remove

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
'
' First original doc
'orDoc.Activate
'
' preprocess for alignment & sorting: remove pictures, convert tables, solve some char problems
'Call Master_Process_Doc_For_Alignment(orDoc)
'

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
OriginalDocument_Order_SortingKeys = Get_MainIDs_Original_SortingOrder_forAllBookmarks(orDoc, "Fado_", nrTopTables - 1) ' top bookmarks will always be ONE extra comp to main bookmarks: with 2 tops, 1 main (theres a top at the end too)


'***************************************************************************************************************************************
'*********************************************** SWITCH TO TARGET DOC PROCESSING *******************************************************
'***************************************************************************************************************************************


' switch active doc for preprocessing
'tgDoc.Activate

'Call Master_Process_Doc_For_Alignment(tgDoc)

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


Sub GetAll_MainIDs_ToNewDocument_Name()      ' GET BOOKMARKS IN NAME ORDER, POSITION SET ASIDE (FOR SITUATION WHEN THEY ARE ARRANGED IN NAME-ORDER)
' More complete verions, extract chapter name, numeric ID & bookmark name

Dim tbk As Bookmark
Dim tDoc As Document
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
                Set tDoc = Documents.Add
                cdoc.Activate
                
                tDoc.Content.InsertAfter ("Bkm " & tbk.Name & "   =   " & _
                    Get_First_NonEmpty_NonNumeric_Paragraph(cdoc, tbk) & "   " & vbTab & "-   ID " & Get_MainID_fromBookmark(cdoc, tbk) & vbCr)
                counter = counter + 1
            Else
                tDoc.Content.InsertAfter ("Bkm " & tbk.Name & "   =   " & _
                    Get_First_NonEmpty_NonNumeric_Paragraph(cdoc, tbk) & "   " & vbTab & "-   ID " & Get_MainID_fromBookmark(cdoc, tbk) & vbCr)
                counter = counter + 1
            End If
        End If
        
    Next j
End If

If counter > 0 Then
    MsgBox "S-au gasit si extras " & counter & " main IDs"
End If

End Sub


' Main routine to check cross-references, it is similar to Extract_AndList_AllHyperlinks, but it does not present the result in a comparative
' table as the other one, where user will have to manually look for eroneous cross-refs.

' This one does same, but for 2 opened documents, original and translation (source and target languages) and eliminates by itself unneeded results...
' presenting to user only useful results: the wrong hyperlinks !
Sub Extract_andCheck_All_CrossReferences()

Dim sDoc As Document, tDoc As Document

If Documents.Count <> 2 Then
    MsgBox "Please open only 2 documents for hyperlink data comparison, source and target languages!"
    Exit Sub
End If


If Documents(1).Name Like "*.en##*" Then
    Set sDoc = Documents(1)
    Set tDoc = Documents(2)
Else
    Set sDoc = Documents(2)
    Set tDoc = Documents(1)
End If


If MsgBox("Shall we proceed with source document as " & vbCr & vbCr & sDoc.Name & vbCr & vbCr & "   and target as " & vbCr & vbCr & tDoc & " ?", vbYesNo + vbQuestion) = vbNo Then
    StatusBar = "Extract_abdCheck_All_CrossReferences: Wrong documents, Quitting..."
    Exit Sub
End If



Dim sourceIDs_and_Hypers
Dim targetIDs_and_Hypers

sourceIDs_and_Hypers = Get_All_IDs_andHyperlinks_V(sDoc)    ' Vertical version, first dimension is not categories of info, but objects number!
targetIDs_and_Hypers = Get_All_IDs_andHyperlinks_V(tDoc)


If UBound(sourceIDs_and_Hypers, 2) <> UBound(targetIDs_and_Hypers, 2) Then
    MsgBox "ERROR while extracting and comparing hyperlinks in FADO source and target documents: extracted " & UBound(sourceIDs_and_Hypers, 2) & " chapters from " & sDoc.Name & _
        UBound(targetIDs_and_Hypers, 2) & " from " & tDoc.Name & "!" & vbCr & "Please run IDs and Title extraction routines and correct docs to obtain perfect extraction!"
    Exit Sub
End If

Debug.Print "Extracted from " & ActiveDocument.Name & " IDInfos for " & UBound(sourceIDs_and_Hypers, 2) & " Fado bookmarks!"


Dim sSourceIDs_and_Hypers
Dim sTargetIDs_and_Hypers


sSourceIDs_and_Hypers = Sort_2D_Array(sourceIDs_and_Hypers)
sTargetIDs_and_Hypers = Sort_2D_Array(targetIDs_and_Hypers)

' The other function also works, call it as such to sort for column 0 (first)
'Call QuickSortArray(sourceIDs_and_Hypers, , , 0)
'Call QuickSortArray(targetIDs_and_Hypers, , , 0)

'sSourceIDs_and_Hypers = sourceIDs_and_Hypers
'sTargetIDs_and_Hypers = targetIDs_and_Hypers

' Not needed anymore, take up memory space
ReDim sourceIDs_and_Hypers(0, 0)
ReDim targetIDs_and_Hypers(0, 0)

Dim tobeCheckedChapters() As String
Dim k As Integer
Dim wrongIDChapters As String


For i = 0 To UBound(sSourceIDs_and_Hypers)
    
    'Debug.Assert (i < 29)
    
    If sSourceIDs_and_Hypers(i, 0) = sTargetIDs_and_Hypers(i, 0) Then
        
        If sSourceIDs_and_Hypers(i, 1) <> sTargetIDs_and_Hypers(i, 1) Or _
           Get_SeparatedString_Differences(sTargetIDs_and_Hypers(i, 2), sSourceIDs_and_Hypers(i, 2)) <> "" Then
            
            k = k + 1
            
            ReDim Preserve tobeCheckedChapters(2, k - 1)
            
            tobeCheckedChapters(0, k - 1) = sSourceIDs_and_Hypers(i, 0)
            tobeCheckedChapters(1, k - 1) = sTargetIDs_and_Hypers(i, 1) - sSourceIDs_and_Hypers(i, 1)
            
            If Get_SeparatedString_Differences(sTargetIDs_and_Hypers(i, 2), sSourceIDs_and_Hypers(i, 2)) <> "" Then
                tobeCheckedChapters(2, k - 1) = Get_SeparatedString_Differences(sTargetIDs_and_Hypers(i, 2), sSourceIDs_and_Hypers(i, 2))
            End If
        
        End If
        
        
    Else
        wrongIDChapters = IIf(wrongIDChapters = "", sTargetIDs_and_Hypers(i, 0) & "(" & sSourceIDs_and_Hypers(i, 0) & ")", _
            wrongIDChapters & "," & sTargetIDs_and_Hypers(i, 0) & "(" & sSourceIDs_and_Hypers(i, 0) & ")")
    End If
    
Next i


If wrongIDChapters <> "" Then
    MsgBox "We have identified following wrongly ID'd chapters: " & wrongIDChapters
End If


If k <> 0 Then
    MsgBox "Found " & k & " chapters with hyperlinks differences in numbers and or target content!"
    Call Create_ComparativeDocument(tobeCheckedChapters)
Else
    MsgBox "Found NO hyperlink differences in numbers or target content in ALL chapters!"
End If



'If UBound(IDs_and_Hypers, 2) > 0 Then
'
'    Documents.Add
'
'    For i = 0 To UBound(IDs_and_Hypers, 2)
'
'        ActiveDocument.StoryRanges(wdMainTextStory).InsertAfter (IDs_and_Hypers(0, i) & vbTab & IDs_and_Hypers(1, i) & vbTab & IDs_and_Hypers(2, i) & vbCr)
'
'    Next i
'
'End If


End Sub


Function Get_SeparatedString_Differences(SourceString1, SourceString2) As String
' Hard-coded pipe separator for now ("|")
' Source1 string parameter is always target (the one we need checked against the source)
' Source2 is always source (the supposedly correct one... the model)

If InStr(1, SourceString1, "|") > 0 And InStr(1, SourceString2, "|") > 0 Then
    
    Dim str1Arr
    Dim str2Arr
    
    str1Arr = Split(SourceString1, "|")     ' TARGET
    str2Arr = Split(SourceString2, "|")     ' SOURCE
    
    If UBound(str1Arr) <> UBound(str2Arr) Then
        
        'Get_SeparatedString_Differences = ""
        'Debug.Print "Get_SeparatedString_Differences: Input string parameters supplied to function have different length - please check !"
        'Exit Function
        Dim difHypNum As Boolean
        difHypNum = True       ' logical flag, represents diff hyp num in source-target
        
    End If
    
    
    Dim tmpStrResult As String
        
    ' Decide max upper limit, which array is smaller... unsafe to go otherwise!
    If difHypNum Then
        
       Dim maxU As Integer
       maxU = IIf(UBound(str1Arr) < UBound(str2Arr), UBound(str1Arr), UBound(str2Arr))
        
    Else
        
        maxU = UBound(str1Arr)
        
    End If
    
    
    ' Now compare and gather only incorrect hyperlinks sub-addresses!
    For i = 0 To maxU
    
        ' str1Arr = target, str2Arr = source
        If (str1Arr(i) <> str2Arr(i)) And Not (str2Arr(i) & "_1" <> str1Arr(i) Or str1Arr(i) & "_1" <> str2Arr(i)) Then
        
            tmpStrResult = IIf(tmpStrResult = "", str1Arr(i) & "[" & str2Arr(i) & "]", tmpStrResult & " | " & str1Arr(i) & "[" & str2Arr(i) & "]")      ' Unified format, we signal to the user the wrong string (str1Arr(i)) against the correct one, between "[]"s - strArr2(i)
        
        End If
    
    Next i
    
    ' we also need go through the rest of the longer array, if number differs, to signal the extra/ missing elements!
    If difHypNum Then
        
        ' Arr1 smaller than Arr2
        If UBound(str1Arr) < UBound(str2Arr) Then
            
            For j = UBound(str1Arr) + 1 To UBound(str2Arr)
                ' Insert extra elements in source
                tmpStrResult = IIf(tmpStrResult = "", Chr(215) & "[" & str2Arr(j) & "]", tmpStrResult & " | " & Chr(215) & "[" & str2Arr(j) & "]")     ' x multiplication sign stands for nothing
            Next j
            
        Else    ' Inverse, Arr2 is smaller than Arr1
            
            ' Insert extra elements in target !
            For j = UBound(str2Arr) + 1 To UBound(str1Arr)
                tmpStrResult = IIf(tmpStrResult = "", str1Arr(j) & "[]", tmpStrResult & " | " & str1Arr(j) & "[]")
            Next j
            
        End If
        
    End If
    
    
Else
    Get_SeparatedString_Differences = ""
    Exit Function
End If

Get_SeparatedString_Differences = tmpStrResult


End Function


Sub Create_ComparativeDocument(Content)         ' Content is actually a bi-dimensional array, first dimension being data categories (table columns)

' Sanity checks
If Not IsArray(Content) Then StatusBar = "Create_ComparativeDocument: Supplied parameter NOT ARRAY!": Exit Sub
' Supplied content array needs first dimension to be three-fold: mainID, number of hypers, hyper sub-addresses content differences ! (these are the data catefories, the tables' columns - headers even)
If Not UBound(Content) = 2 Then StatusBar = "Create_ComparativeDocument: Supplied content array parameter NOT 3 arms!": Exit Sub



Dim resultDoc As Document

Set resultDoc = Documents.Add


Dim resultTable As Table

Set resultTable = resultDoc.Tables.Add(resultDoc.StoryRanges(wdMainTextStory), 1, 3)


' widen table
resultTable.PreferredWidthType = wdPreferredWidthPercent: resultTable.PreferredWidth = 100
resultTable.Columns(1).PreferredWidthType = wdPreferredWidthPercent: resultTable.Columns(2).PreferredWidthType = wdPreferredWidthPercent: resultTable.Columns(3).PreferredWidthType = wdPreferredWidthPercent
resultTable.Columns(1).PreferredWidth = 13: resultTable.Columns(2).PreferredWidth = 22: resultTable.Columns(3).PreferredWidth = 65


resultTable.Rows(1).Cells(1).Range.Text = "Main ID"
resultTable.Rows(1).Cells(2).Range.Text = "Hyperlinks Count Difference"
resultTable.Rows(1).Cells(3).Range.Text = "Hyperlink targets differences (source in [])"


For i = 0 To UBound(Content, 2)
    
    resultTable.Rows.Add
    
    resultTable.Rows(i + 2).Cells(1).Range.Text = Content(0, i)
    resultTable.Rows(i + 2).Cells(2).Range.Text = Content(1, i)
    resultTable.Rows(i + 2).Cells(3).Range.Text = Content(2, i)
    
Next i

With resultTable

    .Borders.InsideLineStyle = wdLineStyleSingle: .Borders.OutsideLineStyle = wdLineStyleSingle
    .Borders.InsideLineWidth = wdLineWidth025pt: .Borders.OutsideLineWidth = wdLineWidth025pt
    .Borders.InsideColorIndex = wdGray25: .Borders.OutsideColorIndex = wdGray25


    .Borders(wdBorderHorizontal).Visible = True: .Borders(wdBorderVertical).Visible = False
    .Borders(wdBorderLeft).Visible = False: .Borders(wdBorderRight).Visible = False
    .Borders(wdBorderTop).Visible = True: .Borders(wdBorderBottom).Visible = True

End With

End Sub


' Identical to but does not present results in a Word doc table, returns bidimensional array instead
Function Get_All_IDs_andHyperlinks(TargetDocument As Document) As String()

Dim tbk As Bookmark
Dim counter As Integer, bcounter As Integer, idx As Integer


Dim sDoc As Document
Set sDoc = TargetDocument

Dim thy As Hyperlink
Dim ttb As Table


Dim tmpResults() As String

If sDoc.Bookmarks.Count > 0 Then

    For Each tbk In sDoc.StoryRanges(wdMainTextStory).Bookmarks   ' accessing by range provides for ordered access
        
        If Left$(tbk.Name, 5) = "Fado_" Then    ' Process only the good hyperlinks
            
            ' Om first FADO bookmark, initialize everything
            If counter = 0 Then
                
                idx = idx + 1
                
                ReDim tmpResults(2, 0)  ' tri-dimensional array returned, first dimension is main ID, second is Hyperlink number, third hyperlinks sub-address info (pipe-separated string)
                
                tmpResults(0, 0) = Get_MainID_fromBookmark(sDoc, tbk)
                ' DONT return chapter title in this ? !? Useful ?
                'Get_First_NonEmpty_NonNumeric_Paragraph(sdoc, tbk)
                            
                If tbk.Range.Hyperlinks.Count > 0 Then
                    
                    counter = counter + 1
                    
                    ' store current bookmarks hyperlinks number in second dimension of temp array
                    tmpResults(1, 0) = tbk.Range.Hyperlinks.Count
                    
                    Dim tmpHypSubInfo As String
                    
                    ' take into account normal hypers also, with web or mail address instead of sub-address !
                    
                    
                    For Each thy In tbk.Range.Hyperlinks
                        If thy.SubAddress <> "" Then    ' cross-reference
                            tmpHypSubInfo = IIf(tmpHypSubInfo = "", thy.SubAddress, tmpHypSubInfo & "|" & thy.SubAddress)   ' make a pipe-separated string with all sub-addresses of all hyps in range
                        Else    ' normal hyperlink, web or mail
                            tmpHypSubInfo = IIf(tmpHypSubInfo = "", thy.Address, tmpHypSubInfo & "|" & thy.Address)
                        End If
                    Next thy
                    
                    
                    tmpResults(2, 0) = tmpHypSubInfo
                    tmpHypSubInfo = ""
                    
                Else    ' No hyperlinks subaddresses to gather, there are none !
                    bcounter = bcounter + 1
                    tmpResults(1, 0) = tbk.Range.Hyperlinks.Count
                End If
                
            Else    ' Second and all other passes
            
                idx = idx + 1
                
                ReDim Preserve tmpResults(2, idx - 1)
                'ttb.Rows.Add
                
                tmpResults(0, idx - 1) = Get_MainID_fromBookmark(sDoc, tbk)
                
                If tbk.Range.Hyperlinks.Count > 0 Then
                        
                    counter = counter + 1
                    
                    tmpResults(1, idx - 1) = tbk.Range.Hyperlinks.Count
                    
                    For Each thy In tbk.Range.Hyperlinks
                        If thy.SubAddress <> "" Then    ' cross-reference
                            tmpHypSubInfo = IIf(tmpHypSubInfo = "", thy.SubAddress, tmpHypSubInfo & "|" & thy.SubAddress)   ' make a pipe-separated string with all sub-addresses of all hyps in range
                        Else    ' normal hyperlink, web or mail
                            tmpHypSubInfo = IIf(tmpHypSubInfo = "", thy.Address, tmpHypSubInfo & "|" & thy.Address)
                        End If
                    Next thy
                    
                    tmpResults(2, idx - 1) = tmpHypSubInfo
                    tmpHypSubInfo = ""
                                                    
                Else    ' no hypers in current bkm
                
                    bcounter = bcounter + 1
                    tmpResults(1, idx - 1) = tbk.Range.Hyperlinks.Count
                    
                End If
            
            End If
        
        End If  ' If tbk.name este "Fado_###"
        
    Next tbk

End If


Get_All_IDs_andHyperlinks = tmpResults

End Function

' Identical to but does not present results in a Word doc table, returns bidimensional array instead
' VERTICAL version of the above, main dimension is NOT number of categories of information, but actual "objects" number...
' CHAPTERS info in this case !
Function Get_All_IDs_andHyperlinks_V(TargetDocument As Document) As String()


Dim tbk As Bookmark
Dim counter As Integer, bcounter As Integer, idx As Integer


Dim sDoc As Document
Set sDoc = TargetDocument

Dim thy As Hyperlink
Dim ttb As Table


Dim chaptersNumber
chaptersNumber = Count_TopTables(TargetDocument) - 1    ' TOP tables are always one more than bookmarks


Dim tmpResults() As String


If sDoc.Bookmarks.Count > 0 Then

    For Each tbk In sDoc.StoryRanges(wdMainTextStory).Bookmarks   ' accessing by range provides for ordered access
        
        If Left$(tbk.Name, 5) = "Fado_" Then    ' Process only the good hyperlinks
            
            ' Om first FADO bookmark, initialize everything
            If counter = 0 Then
                
                ' counting FADO bookmarks
                idx = idx + 1
                
                ' bi-dimensional array returned, first dimension is actual chapters themselves, second dimension represents info saught for each chapter, namely
                ' main ID, Hyperlink number and hyperlinks sub-address info (pipe-separated string)
                ReDim tmpResults(chaptersNumber - 1, 2)
                
                
                tmpResults(0, 0) = Get_MainID_fromBookmark(sDoc, tbk)
                ' DONT return chapter title in this ? !? Useful ?
                'Get_First_NonEmpty_NonNumeric_Paragraph(sdoc, tbk)
                            
                If tbk.Range.Hyperlinks.Count > 0 Then
                    
                    ' Counting FADO bookmarks with hypers
                    counter = counter + 1
                    
                    ' store current bookmarks hyperlinks number in second dimension of temp array
                    tmpResults(0, 1) = tbk.Range.Hyperlinks.Count
                    
                    
                    Dim tmpHypSubInfo As String
                    
                    ' take into account normal hypers also, with web or mail address instead of sub-address !
                    For Each thy In tbk.Range.Hyperlinks
                        If thy.SubAddress <> "" Then    ' cross-reference
                            tmpHypSubInfo = IIf(tmpHypSubInfo = "", thy.SubAddress, tmpHypSubInfo & "|" & thy.SubAddress)   ' make a pipe-separated string with all sub-addresses of all hyps in range
                        Else    ' normal hyperlink, web or mail
                            tmpHypSubInfo = IIf(tmpHypSubInfo = "", thy.Address, tmpHypSubInfo & "|" & thy.Address)
                        End If
                    Next thy
                    
                    
                    tmpResults(0, 2) = tmpHypSubInfo
                    tmpHypSubInfo = ""
                    
                    
                Else    ' No hyperlinks subaddresses to gather, there are none !
                    bcounter = bcounter + 1
                    tmpResults(0, 1) = tbk.Range.Hyperlinks.Count
                End If
                
            Else    ' Second and all other passes
                
                ' Counting FADO bookmarks
                idx = idx + 1
                
                ' Array already dimensioned, no need to add extra "row" all the time !
                'ReDim Preserve tmpResults(idx - 1, 2)
                                
                tmpResults(idx - 1, 0) = Get_MainID_fromBookmark(sDoc, tbk)
                
                If tbk.Range.Hyperlinks.Count > 0 Then
                        
                    ' Counting FADO bkms with hypers
                    counter = counter + 1
                    
                    tmpResults(idx - 1, 1) = tbk.Range.Hyperlinks.Count
                    
                    For Each thy In tbk.Range.Hyperlinks
                        If thy.SubAddress <> "" Then    ' cross-reference
                            tmpHypSubInfo = IIf(tmpHypSubInfo = "", thy.SubAddress, tmpHypSubInfo & "|" & thy.SubAddress)   ' make a pipe-separated string with all sub-addresses of all hyps in range
                        Else    ' normal hyperlink, web or mail
                            tmpHypSubInfo = IIf(tmpHypSubInfo = "", thy.Address, tmpHypSubInfo & "|" & thy.Address)
                        End If
                    Next thy
                    
                    tmpResults(idx - 1, 2) = tmpHypSubInfo
                    tmpHypSubInfo = ""
                                                    
                Else    ' no hypers in current bkm
                
                    bcounter = bcounter + 1
                    tmpResults(idx - 1, 1) = tbk.Range.Hyperlinks.Count
                    
                End If
            
            End If
        
        End If  ' If tbk.name este "Fado_###"
        
    Next tbk

End If


Get_All_IDs_andHyperlinks_V = tmpResults

End Function


Sub Extract_AndList_AllHyperlinks()
' Put into new document all fields of type hyperlink, list their sub-address

Dim tbk As Bookmark
Dim counter As Integer, bcounter As Integer, idx As Integer

Dim sDoc As Document, tDoc As Document
Set sDoc = ActiveDocument

Dim thy As Hyperlink
Dim ttb As Table

Dim tsubs As String

If ActiveDocument.Bookmarks.Count > 0 Then
    For Each tbk In ActiveDocument.StoryRanges(wdMainTextStory).Bookmarks   ' accessing by range provides for ordered access
        
        If Left$(tbk.Name, 5) = "Fado_" Then    ' Process only the good hyperlinks
        
            If counter = 0 Then
                idx = idx + 1
                Set tDoc = Documents.Add: tDoc.Tables.Add tDoc.Paragraphs(1).Range, 1, 3: Set ttb = tDoc.Tables(1): sDoc.Activate
                
                ttb.Rows(1).Cells(1).Range.Text = "Main ID": ttb.Rows(1).Cells(2).Range.Text = "Chapter title": ttb.Rows(1).Cells(3).Range.Text = "Hyperlinks and targets"
                
                ttb.Rows.Add
                ttb.Rows(2).Cells(1).Range.Text = Get_MainID_fromBookmark(sDoc, tbk)
                ttb.Rows(2).Cells(2).Range.Text = Get_First_NonEmpty_NonNumeric_Paragraph(sDoc, tbk)
                            
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
                
                ttb.Rows(ttb.Rows.Count).Cells(1).Range.Text = Get_MainID_fromBookmark(sDoc, tbk)
                ttb.Rows(ttb.Rows.Count).Cells(2).Range.Text = Get_First_NonEmpty_NonNumeric_Paragraph(sDoc, tbk)
                            
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

tDoc.SpellingChecked = True    ' Get rid of pesky red underlining

End If


If counter > 0 Then
    If bcounter = 0 Then
        MsgBox "S-au procesat " & counter & " bookmarkuri cu cel putin 1 hiperlinkuri"
    Else
        MsgBox "Au fost gasite " & bcounter & " bookmarkuri fara hiperlinkuri"
    End If
    
End If

End Sub


' Complementary check to Extract_AndList_AllHyperlinks check, in which we check the actual location of all ID bookmarks
Sub All_IDs_Bookmarks_LocationCheck_CtDoc()

Dim tmpResult() As String

tmpResult = Get_IDs_Bookmarks_LocationCheck(ActiveDocument)


If Not IsNotDimensionedArr(tmpResult) Then
    
    Dim bad_IDBkms_Names As String
    
    bad_IDBkms_Names = Join(tmpResult, " | ")
    
    MsgBox "FOUND " & UBound(tmpResult) + 1 & " badly located ID bookmarks:" & vbCr & vbCr & _
        bad_IDBkms_Names
    
Else
    MsgBox "All ID Bookmarks location check OK !"
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
            
            Selection.InsertAfter Get_First_NonEmpty_NonNumeric_Paragraph(ActiveDocument, tbk)     ' To create hyperlink, insert text first
            ActiveDocument.Hyperlinks.Add Selection.Range, , "_" & _
                Replace(Get_MainID_fromBookmark(ActiveDocument, tbk), Chr(160), ""), , Selection.Text
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


Function Check_andFix_allIDs(TargetDocument As Document) As Boolean


Dim bkm As Bookmark

If TargetDocument.Bookmarks.Count > 0 Then
    
    For Each bkm In TargetDocument.Bookmarks
        If Left(bkm.Name, 5) = "Fado_" Then
            Call Repare_IDs_forBookmark(bkm)
        End If
    Next bkm
    
Else
    Check_andFix_allIDs = False
End If

Check_andFix_allIDs = True


End Function


Sub Repare_IDs_forBookmark(TargetBookmark As Bookmark)

If TargetBookmark.Range.Tables.Count > 0 Then
    
    Dim tmpRange As Range
    
    Set tmpRange = TargetBookmark.Range.Tables(1).Range
    tmpRange.Collapse (wdCollapseEnd)
    tmpRange.SetRange tmpRange.Start, TargetBookmark.Range.End
    
    Dim colIDParagraphs As Collection
    
    colIDParagraphs = Get_IDsParagraphs_forRange(tmpRange)
    
    Call Repair_IDs_fromCollection(TargetBookmark, colIDParagraphs)
    
Else

End If

End Sub


Sub Repair_IDs_fromCollection(TargetBookmark As Bookmark, TargetParagraphsCollection As Collection)


Dim tmpPar As Paragraph


If TargetParagraphsCollection.Count = 1 Then
    
    Set tmpPar = TargetParagraphsCollection.Item(1)
    Call Repair_MainIDParagraph(tmpPar)
    
ElseIf TargetParagraphsCollection.Count > 1 Then
    
    For i = 1 To TargetParagraphsCollection.Count
        Set tmpPar = TargetParagraphsCollection.Item(i)
        Call Repair_MainIDParagraph(tmpPar)
    Next i
    
ElseIf TargetParagraphsCollection = 0 Then
    Debug.Print "Repair_IDs: Found NO ID paragraphs for bkm " & TargetBookmark.Name
End If


End Sub


Sub Repair_MainIDParagraph(ByRef TargetParagraph As Paragraph)

' add NB space before if par seems to be main ID one (like ### or like [space]### or maybe even ###,### with or without spaces)
If TargetParagraph.Range.Text Like "###" Or TargetParagraph.Range.Text Like " ###" Then
    TargetParagraph.Range.Text = Chr(160) & TargetParagraph.Range.Text
    Debug.Print "Repair_MainIDParagraphs: Repaired main ID par " & TargetParagraph.Parent.Index
End If


End Sub


Function Get_IDsParagraphs_forRange(TargetRange As Range) As Collection

Dim tmpPar As Paragraph
    
    Dim IDsCollection As Collection
    
    Set tmpPar = TargetRange.Paragraphs(1)
    
    
    Dim tmpParText As String
    tmpParText = tmpPar.Range.Text
    
    ' we clean away parantheses, simple and hard spaces and we should be left with numerics in the case of ID paragraphs
    tmpParText = Replace(Replace(Replace(Replace(tmpParText, "(", ""), ")", ""), " ", ""), Chr(160), "")
    tmpParText = Replace(tmpParText, vbCr, "")
    
    
    ' we cleaned spaces and parantheses, for id paragraphs now should be left only numbers. We also ignore empty paragraphs
    Do While IsNumeric(tmpParText) Or tmpParText.Text = ""
    
        If IsNumeric(tmpParText) Then
        
            IDsCollection.Add (tmpPar)
        
        End If
    
    Loop
    
    Get_IDsParagraphs_forRange = IDsCollection

End Function


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

Debug.Print "Au fost setate " & topCount & " bookmarks la tabelele TOP! in " & TargetDocument.Name


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


Debug.Print "Au fost setate " & k & " bookmarkuri <Fado_> in " & TargetDocument.Name


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
                
                ' remove temp bookmark, no longer needed
                ctChapter_tmpBkm.Delete
                
                ' "topS" & format(j+1, "000")).
                TargetDocument.Bookmarks.Add "Fado_" & Format(ctChapter_mainID_origPosition + 1, "000"), ctChapter_range    ' +1 because original Fado_ bkms are 1 based (they count themselves)
                
            End If
        
        End If
        
    Next j
    
End If


Debug.Print "Au fost setate " & k & " bookmarkuri <Fado_> in " & TargetDocument.Name



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
            
            ReDim Preserve tmpKeys(k - 1)
            
            tfn_par = Get_MainID_fromBookmark(TargetDocument, ck_tbk)
            
            If tfn_par <> "" Then
                tmpKeys(k - 1) = Replace(tfn_par, Chr(160), "")
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
frmSortPages.Show

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
Dim sDoc As Document

Set sDoc = ActiveDocument
Set tgDoc = Documents.Add
sDoc.Activate

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
Dim sDoc As Document

Set sDoc = ActiveDocument
Set tgDoc = Documents.Add
sDoc.Activate

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
        
        ctIS.Select
        
        ctIS.ConvertToShape
        
        'ctIS.Type = wdInlineShapePicture
        
    End If
Next ctIS

MsgBox "Converted " & counter & " inline embedded OLEobjects to inline shapes!", vbOKOnly + vbInformation, "Done"



End Sub

Sub Highlight_All_TopTables_CvTxt_Rest(TargetDocument As Document)

Dim tt As Table


If TargetDocument.Tables.Count > 0 Then
    
    Dim topTablesCount As Integer
    Dim mrTablesCount As Integer
    
    For Each tt In TargetDocument.Tables
        
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
    
    Debug.Print "Highlighted " & topTablesCount & " top tables, highlighted & converted txt " & mrTablesCount & " multiple rows tables in " & TargetDocument.Name
    
Else
    Debug.Print "Found NO tables in " & TargetDocument.Name
End If


End Sub

Function Count_TopTables_CtDoc() As Integer

Dim tt As Table

Dim topCount As Integer


For Each tt In ActiveDocument.Tables
    
    If tt.Rows.Count = 1 Then
        
        If tt.Rows(1).Cells.Count = 1 Then
            
            If tt.Shading.BackgroundPatternColorIndex = 16 Then
                
                'If tt.Borders.OutsideLineStyle = 24 Then
                'tt.Rows(1).Cells(1).Range.Text Like "??top??" Or _
                tt.Rows(1).Cells(1).Range.Text Like "???nceput??" Or _
                tt.Rows(1).Cells(1).Range.Text Like "??nceput??" Then
                    topCount = topCount + 1
                ' Our "top" tables !
                'tt.Range.Shading.BackgroundPatternColorIndex = wdYellow
                
'                Set tempRange = tt.Range
'                tempRange.Collapse (wdCollapseStart)
'                tempRange.MoveEnd wdCharacter, -1
'                TargetDocument.Bookmarks.Add "topS" & Format(topCount, "000"), tempRange
                'End If
                
            End If
            
        End If
        
    End If
    
Next tt

Count_TopTables_CtDoc = topCount

Debug.Print "Numarat " & topCount & " tabele TOP in " & ActiveDocument.Name

End Function


Sub Green_All_GoodTopTables()

Dim tt As Table

Dim topCount As Integer


For Each tt In ActiveDocument.Tables
    
    If tt.Rows.Count = 1 Then
        
        If tt.Rows(1).Cells.Count = 1 Then
            
            If tt.Shading.BackgroundPatternColorIndex = 16 Then
                
                If tt.Borders.OutsideLineStyle = 24 Then
                'tt.Rows(1).Cells(1).Range.Text Like "??top??" Or _
                tt.Rows(1).Cells(1).Range.Text Like "???nceput??" Or _
                tt.Rows(1).Cells(1).Range.Text Like "??nceput??" Then
                
                
                topCount = topCount + 1
                tt.Shading.BackgroundPatternColorIndex = wdBrightGreen
                
                ' Our "top" tables !
                'tt.Range.Shading.BackgroundPatternColorIndex = wdYellow
                
'                Set tempRange = tt.Range
'                tempRange.Collapse (wdCollapseStart)
'                tempRange.MoveEnd wdCharacter, -1
'                TargetDocument.Bookmarks.Add "topS" & Format(topCount, "000"), tempRange
                End If
                
            End If
            
        End If
        
    End If
    
Next tt



Debug.Print "Numarat " & topCount & " tabele TOP in " & ActiveDocument.Name


End Sub


Sub Gray_All_GoodTopTables()
' BACK to original colour

Dim tt As Table

Dim topCount As Integer


For Each tt In ActiveDocument.Tables
    
    If tt.Rows.Count = 1 Then
        
        If tt.Rows(1).Cells.Count = 1 Then
            
            If tt.Shading.BackgroundPatternColorIndex = 4 Then      ' GREEN
                
                If tt.Borders.OutsideLineStyle = 24 Then
                'tt.Rows(1).Cells(1).Range.Text Like "??top??" Or _
                tt.Rows(1).Cells(1).Range.Text Like "???nceput??" Or _
                tt.Rows(1).Cells(1).Range.Text Like "??nceput??" Then
                
                
                topCount = topCount + 1
                tt.Shading.BackgroundPatternColorIndex = 16
                
                ' Our "top" tables !
                'tt.Range.Shading.BackgroundPatternColorIndex = wdYellow
                
'                Set tempRange = tt.Range
'                tempRange.Collapse (wdCollapseStart)
'                tempRange.MoveEnd wdCharacter, -1
'                TargetDocument.Bookmarks.Add "topS" & Format(topCount, "000"), tempRange
                End If
                
            End If
            
        End If
        
    End If
    
Next tt



Debug.Print "Numarat " & topCount & " tabele TOP in " & ActiveDocument.Name


End Sub

Sub Identify_First_MultiRow_TopTable()

Dim tt As Table

Dim topCount As Integer


For Each tt In ActiveDocument.Tables
    
    'If tt.Borders.OutsideLineStyle = 24 Then
    
        If tt.Rows.Count = 1 Then
        
        'If tt.Rows(1).Cells.Count = 1 Then
            
            If tt.Shading.BackgroundPatternColorIndex <> wdAuto And _
               tt.Shading.BackgroundPatternColorIndex <> 16 Then
                
                
                tt.Select
                Exit For
                
                'tt.Rows(1).Cells(1).Range.Text Like "??top??" Or _
                tt.Rows(1).Cells(1).Range.Text Like "???nceput??" Or _
                tt.Rows(1).Cells(1).Range.Text Like "??nceput??" Then
                    
                'topCount = topCount + 1
                ' Our "top" tables !
                'tt.Range.Shading.BackgroundPatternColorIndex = wdYellow
                
'                Set tempRange = tt.Range
'                tempRange.Collapse (wdCollapseStart)
'                tempRange.MoveEnd wdCharacter, -1
'                TargetDocument.Bookmarks.Add "topS" & Format(topCount, "000"), tempRange
                'End If
                
            End If
            
        End If
        
    'End If
    
Next tt



Debug.Print "Numarat " & topCount & " tabele TOP in " & ActiveDocument.Name

End Sub



Function Count_TopTables(TargetDocument As Document) As Integer


Dim tt As Table

Dim topCount As Integer


For Each tt In TargetDocument.Tables
    If tt.Rows.Count = 1 Then
        If tt.Rows(1).Cells.Count = 1 Then
            If tt.Shading.BackgroundPatternColorIndex = 16 Then
                
                'If tt.Rows(1).Cells(1).Range.Text Like "??top??" Then
                
                'tt.Rows(1).Cells(1).Range.Text Like "??top??" Or _
                tt.Rows(1).Cells(1).Range.Text Like "???nceput??" Or _
                tt.Rows(1).Cells(1).Range.Text Like "??nceput??" Then
                    topCount = topCount + 1
                ' Our "top" tables !
                'tt.Range.Shading.BackgroundPatternColorIndex = wdYellow
                
'                Set tempRange = tt.Range
'                tempRange.Collapse (wdCollapseStart)
'                tempRange.MoveEnd wdCharacter, -1
'                TargetDocument.Bookmarks.Add "topS" & Format(topCount, "000"), tempRange
                'End If
            End If
        End If
    End If
Next tt

Count_TopTables = topCount

Debug.Print "Numarat " & topCount & " tabele TOP in " & TargetDocument.Name


End Function

Sub Select_BadColorIndex_TopTable()


Dim tt As Table

Dim topCount As Integer


For Each tt In ActiveDocument.Tables
    If tt.Rows.Count = 1 Then
        If tt.Rows(1).Cells.Count = 1 Then
            If tt.Shading.BackgroundPatternColorIndex <> 16 Then
                'tt.Rows(1).Cells(1).Range.Text Like "??top??" Or _
                tt.Rows(1).Cells(1).Range.Text Like "???nceput??" Or _
                tt.Rows(1).Cells(1).Range.Text Like "??nceput??" Then
                'topCount = topCount + 1
                
                tt.Range.Select
                Exit For
                
                ' Our "top" tables !
                'tt.Range.Shading.BackgroundPatternColorIndex = wdYellow
                
'                Set tempRange = tt.Range
'                tempRange.Collapse (wdCollapseStart)
'                tempRange.MoveEnd wdCharacter, -1
'                TargetDocument.Bookmarks.Add "topS" & Format(topCount, "000"), tempRange
            End If
        End If
    End If
Next tt


Debug.Print "Bad color index top table 1 selected"


End Sub

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
realtext = Replace(Replace(Replace(realtext, vbCrCr, ""), " ", ""), Chr(160), "")
realtext = Replace(Replace(realtext, Chr(13), ""), Chr(7), "")

If realtext <> "" Then
    IsTable_Empty = False
Else
    IsTable_Empty = True
End If
        
End Function

Sub AllVBCr_Pairs_ToVbCr(TargetDocument As Document)
' Clean empty paragraphs preceded by other "space symbols"
' such as Tab+Enter, Space+Enter, HardSpace+Enter

Dim ls As String
ls = gscrolib.GetListSeparatorFromRegistry

If ls = "" Then MsgBox "List separator not extracted from registry!", vbOKOnly, "Please debug !": Exit Sub

With TargetDocument.StoryRanges(wdMainTextStory).Find
    .ClearAllFuzzyOptions
    .ClearFormatting
    .Format = False
    .Forward = True
    .Wrap = wdFindStop
        
    .Execute FindText:="([ ^t^s]{1" & ls & "})(^13)", MatchWildcards:=True, Forward:=True, Wrap:=wdFindStop, Format:=False, _
        ReplaceWith:="^p", Replace:=wdReplaceAll
End With


End Sub

Sub AllMultiple_VBCrs_To_Single(TargetDocument As Document)

Dim ls As String
ls = gscrolib.GetListSeparatorFromRegistry

If ls = "" Then MsgBox "List separator not extracted from registry!", vbOKOnly, "Please debug !": Exit Sub

With TargetDocument.StoryRanges(wdMainTextStory).Find
    .ClearAllFuzzyOptions
    .ClearFormatting
    .Format = False
    .Forward = True
    .Wrap = wdFindStop
        
    .Execute FindText:="(^13{2" & ls & "})", MatchWildcards:=True, Forward:=True, Wrap:=wdFindStop, Format:=False, _
        ReplaceWith:="^p", Replace:=wdReplaceAll
End With


End Sub



Function Get_First_NonEmpty_NonNumeric_Paragraph(TargetDocument As Document, TargetBookmark As Bookmark) As String

Dim tpar As Paragraph
Dim realtext As String

If TargetDocument.Bookmarks.Count > 0 Then
    If TargetDocument.Bookmarks.Exists(TargetBookmark) Then
        For Each tpar In TargetBookmark.Range.Paragraphs
            
            If Not tpar.Range.Information(wdWithInTable) Then
                ' Trying to remove all expectable non-printing characters, to retrieve text only as it were
                realtext = tpar.Range.Text
                realtext = Replace(Replace(Replace(realtext, vbCr, ""), vbLf, ""), vbCrLf, "")
                realtext = Replace(Replace(Replace(realtext, Chr(12), ""), Chr(160), " "), "  ", " ")
                realtext = Replace(Replace(Replace(realtext, Chr(10), ""), Chr(13), ""), Chr(7), "")
                realtext = Replace(Replace(realtext, Chr(1), ""), Chr(13), "")
                
                If realtext <> "" And realtext <> Chr(1) & "inceput" And realtext <> " " And _
                   (Not realtext Like "(#*#)" And Not realtext Like " ###*" And _
                   Not realtext Like "?top" And Not realtext Like "?(#*#)" And _
                   Not realtext Like "###") Then
                        Get_First_NonEmpty_NonNumeric_Paragraph = realtext
                        Exit For
                Else
                End If
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
Dim tDoc As Document
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
            Set tDoc = Documents.Add
            ctdoc.Activate
            
            tDoc.Content.InsertAfter (Get_First_NonEmpty_NonNumeric_Paragraph(ctdoc, tbk)) & vbCr
            counter = counter + 1
        Else
            tDoc.Content.InsertAfter (Get_First_NonEmpty_NonNumeric_Paragraph(ctdoc, tbk)) & vbCr
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
            mainID = Replace(Replace(Replace(mainID, vbCr, ""), vbLf, ""), vbCrLf, "")
            mainID = Replace(Replace(mainID, Chr(12), ""), "  ", " ")
            mainID = Replace(Replace(Replace(mainID, Chr(10), ""), Chr(13), ""), Chr(7), "")
            
            If mainID Like Chr(160) & "###*" Then
                tres = mainID
                
                If InStr(1, tres, ",") > 0 Then
                    Get_MainID_fromBookmark = Replace(Split(tres, ",")(0), Chr(160), "")   ' Chapter has actually more than one main ID, separated by comma
                Else
                    Get_MainID_fromBookmark = Replace(tres, Chr(160), "")
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
        MsgBox "Am identificat si stiluit " & counter & " main IDs" & vbCrCr & _
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



Sub AllTables_ConvertToText(TargetDocument As Document)

Dim tbl As Table, tblcount As Integer

If Documents.Count > 0 Then   ' if document is opened
    If TargetDocument.Tables.Count > 0 Then
        For Each tbl In TargetDocument.Tables
            tblcount = tblcount + 1
            tbl.ConvertToText
        Next tbl
        StatusBar = "All " & tblcount & " tables were converted txt!"
    End If
Else
    StatusBar = "Not working without opened document!"
End If


End Sub

Sub AllTables_ConvertToText_CtDoc()

Dim tbl As Table, tblcount As Integer

If Documents.Count > 0 Then   ' if document is opened
    If ActiveDocument.Tables.Count > 0 Then
        For Each tbl In ActiveDocument.Tables
            tblcount = tblcount + 1
            tbl.ConvertToText
        Next tbl
        StatusBar = "All " & tblcount & " tables were converted txt!"
    End If
Else
    StatusBar = "Not working without opened document!"
End If


End Sub


Sub AllInlineImagesDelete(TargetDocument As Document)

Dim ishape As InlineShape

If Documents.Count > 0 Then
    If TargetDocument.InlineShapes.Count > 0 Then
        For Each ishape In TargetDocument.InlineShapes
            ishape.Range.Text = ""
        Next ishape
    Else
        StatusBar = "No inline shapes in document " & TargetDocument.Name
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
        Next shp
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
        Call AllInlineImagesDelete(ActiveDocument)
    End If
    
    If ActiveDocument.Shapes.Count > 0 Then
        Call AllShapesDelete
    End If
End If

End Sub

Sub AllPageBreaks_ToEnters(TargetDocument As Document)

With TargetDocument.StoryRanges(wdMainTextStory).Find
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

Sub RemoveAll_PagebreakBefores(TargetDocument As Document)


With TargetDocument.StoryRanges(wdMainTextStory).Find
    
    .ClearAllFuzzyOptions
    .ClearFormatting
    .ClearHitHighlight
    
    .Format = True
    .Forward = True
    
    .Execute FindText:="", ReplaceWith:="", Replace:=wdReplaceAll
    
End With


End Sub


Function Get_All_Hidden_IDBookmarks(TargetDocument As Document) As String()
' Bookmarks "_###" and "_(###)

If TargetDocument.Hyperlinks.Count > 0 Then
       
    TargetDocument.Bookmarks.ShowHidden = True
           
    Dim ctBkm As Bookmark
    
    Dim tmpArray() As String
    
    Dim k As Integer
    Let k = 0
    
    For Each ctBkm In TargetDocument.Bookmarks
        
        If Left(ctBkm.Name, 1) = "_" And IsNumeric(Mid(ctBkm.Name, 2, 1)) Then
            k = k + 1
            ReDim Preserve tmpArray(k - 1)
            tmpArray(UBound(tmpArray)) = ctBkm.Name
        End If
        
    Next ctBkm
    
    Get_All_Hidden_IDBookmarks = tmpArray
    
Else
    Get_All_Hidden_IDBookmarks = Null
End If


End Function


Function Get_IDs_Bookmarks_LocationCheck(TargetDocument As Document) As String()


Dim allID_Bookmarks() As String

allID_Bookmarks = Get_All_Hidden_IDBookmarks(TargetDocument)


If IsNotDimensionedArr(allID_Bookmarks) Then
    MsgBox "Get_IDs_Bookmarks_LocationCheck: Error, could not find ANY hidden <_###> bookmark in document! These should not be missing! Please check existence/ remedy and retry" & vbCr & vbCr & _
        "Also please be informed that sometimes, it suffises to check existence of these bookmarks, show/ hide hidden bookmarks and they will magically appear"
    'Get_IDs_Bookmarks_LocationCheck = Null
    Exit Function
End If


Dim tmpResult() As String

Dim k As Integer
k = 0

For i = 0 To UBound(allID_Bookmarks)
    
    Dim ctBkmRng As Range
    
    Set ctBkmRng = TargetDocument.Bookmarks(allID_Bookmarks(i)).Range
    
    ' Majority of bookmarks will be empty, placed in front of chapter number
    'If ctBkmRng.Characters.Count = 0 Then
    
        ctBkmRng.MoveEndUntil (vbCr)
        
    'End If
    
    
    Dim ctBkmRngText As String
    
    ctBkmRngText = ctBkmRng.Text
    
    ctBkmRngText = Replace(Replace(Replace(ctBkmRngText, "(", ""), ")", ""), Chr(160), "")
    ctBkmRngText = Replace(ctBkmRngText, " ", "")
    
    
    ' need to take into account that bkm name begins with an extra "_" and that some
    ' bkms even have an extra "_1" at the end !
    Dim ctBkmNameEssential As String
    
    ctBkmNameEssential = allID_Bookmarks(i)
    
    
    If Right(ctBkmNameEssential, 2) = "_1" Then
        
        ' multiple "_", beginning and near end
        ctBkmNameEssential = Replace(Left(allID_Bookmarks(i), Len(allID_Bookmarks(i)) - 2), "_", "")
        
        
    Else
        ' one "_" only, beginning
        ctBkmNameEssential = Replace(allID_Bookmarks(i), "_", "")
        
    End If
    
    
    If ctBkmRngText <> ctBkmNameEssential Then
        k = k + 1
        ReDim Preserve tmpResult(k - 1)
        tmpResult(UBound(tmpResult)) = allID_Bookmarks(i)
    End If
    
    
Next i

Get_IDs_Bookmarks_LocationCheck = tmpResult


End Function

Function Sort_2D_Array(ByVal ArrayToSort As Variant, Optional SortDescending, Optional SortBySecondDimension) As Variant
         
    Dim tmpArray    ' to use during work
    tmpArray = ArrayToSort
    
    'Dim tmpArray(5, 2) As Variant
    'Dim v As Variant
    Dim i As Integer, j As Integer
    Dim r As Integer, c As Integer
    Dim temp As Variant
     
     'Create 2-dimensional array
     
'    v = Array(56, 22, "xyz", 22, 30, "zyz", 56, 30, "zxz", 22, 30, "zxz", 10, 18, "zzz", 22, 18, "zxx")
'    For i = 0 To UBound(v)
'        tmpArray(i \ 3, i Mod 3) = v(i)
'    Next
    
    Debug.Print "Unsorted array:"
    For r = LBound(tmpArray) To UBound(tmpArray)
        For c = LBound(tmpArray, 2) To UBound(tmpArray, 2)
            Debug.Print tmpArray(r, c);
        Next
        Debug.Print
    Next
    
    
     'Bubble sort column 0
    
    If IsMissing(SortDescending) Then
        ' Ascending
        For i = LBound(tmpArray) To UBound(tmpArray) - 1
            For j = i + 1 To UBound(tmpArray)
                If tmpArray(i, 0) > tmpArray(j, 0) Then
                    For c = LBound(tmpArray, 2) To UBound(tmpArray, 2)
                        temp = tmpArray(i, c)
                        tmpArray(i, c) = tmpArray(j, c)
                        tmpArray(j, c) = temp
                    Next
                End If
            Next
        Next
    Else     ' Descending
        For i = LBound(tmpArray) To UBound(tmpArray) - 1
            For j = i + 1 To UBound(tmpArray)
                If tmpArray(i, 0) < tmpArray(j, 0) Then
                    For c = LBound(tmpArray, 2) To UBound(tmpArray, 2)
                        temp = tmpArray(i, c)
                        tmpArray(i, c) = tmpArray(j, c)
                        tmpArray(j, c) = temp
                    Next
                End If
            Next
        Next
    End If
     
    
    If Not IsMissing(SortBySecondDimension) Then
        'Bubble sort column 1, where adjacent rows in column 0 are equal
        If IsMissing(SortDescending) Then
            ' Ascending
            For i = LBound(tmpArray) To UBound(tmpArray) - 1
                For j = i + 1 To UBound(tmpArray)
                    If tmpArray(i, 0) = tmpArray(j, 0) Then
                        If tmpArray(i, 1) > tmpArray(j, 1) Then
                            For c = LBound(tmpArray, 2) To UBound(tmpArray, 2)
                                temp = tmpArray(i, c)
                                tmpArray(i, c) = tmpArray(j, c)
                                tmpArray(j, c) = temp
                            Next
                        End If
                    End If
                Next
            Next
         Else    ' Descending
            For i = LBound(tmpArray) To UBound(tmpArray) - 1
                For j = i + 1 To UBound(tmpArray)
                    If tmpArray(i, 0) = tmpArray(j, 0) Then
                        If tmpArray(i, 1) < tmpArray(j, 1) Then
                            For c = LBound(tmpArray, 2) To UBound(tmpArray, 2)
                                temp = tmpArray(i, c)
                                tmpArray(i, c) = tmpArray(j, c)
                                tmpArray(j, c) = temp
                            Next
                        End If
                    End If
                Next
            Next
         End If
    End If
     
     'Output sorted array
'    Debug.Print "Sorted array:"
'    For r = LBound(tmpArray) To UBound(tmpArray)
'        For c = LBound(tmpArray, 2) To UBound(tmpArray, 2)
'            Debug.Print tmpArray(r, c);
'        Next
'        Debug.Print
'    Next
    
    Sort_2D_Array = tmpArray
    
End Function

Public Sub QuickSortArray(ByRef SortArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0)
    On Error Resume Next

    'Sort a 2-Dimensional array

    ' SampleUsage: sort arrData by the contents of column 3
    '
    '   QuickSortArray arrData, , , 3

    '
    'Posted by Jim Rech 10/20/98 Excel.Programming

    'Modifications, Nigel Heffernan:

    '       ' Escape failed comparison with empty variant
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim arrRowTemp As Variant
    Dim lngColTemp As Long

    If IsEmpty(SortArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortArray), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortArray, 1)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortArray, 1)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = SortArray((lngMin + lngMax) \ 2, lngColumn)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(SortArray(n)) - varMid *might* pick up a valid default member or property
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j
        While SortArray(i, lngColumn) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortArray(j, lngColumn) And j > lngMin
            j = j - 1
        Wend

        If i <= j Then
            ' Swap the rows
            ReDim arrRowTemp(LBound(SortArray, 2) To UBound(SortArray, 2))
            For lngColTemp = LBound(SortArray, 2) To UBound(SortArray, 2)
                arrRowTemp(lngColTemp) = SortArray(i, lngColTemp)
                SortArray(i, lngColTemp) = SortArray(j, lngColTemp)
                SortArray(j, lngColTemp) = arrRowTemp(lngColTemp)
            Next lngColTemp
            Erase arrRowTemp

            i = i + 1
            j = j - 1
        End If
    Wend

    If (lngMin < j) Then Call QuickSortArray(SortArray, lngMin, j, lngColumn)
    If (i < lngMax) Then Call QuickSortArray(SortArray, i, lngMax, lngColumn)

End Sub

'
' **********************************************************************************************************
' ***********   Routines for sorting blocks of text, made for ST 6781/12 REV1 & ST 9916/08 *****************
' **********************************************************************************************************
'
