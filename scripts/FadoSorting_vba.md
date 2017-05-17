# VBA Project: **FadoSorting**
## VBA Module: **[FadoSorting](/scripts/FadoSorting.vba "source is here")**
### Type: StdModule  

This procedure list for repo (FadoSorting) was automatically created on 17/05/2017 19:11:08 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in FadoSorting

---
VBA Procedure: **Master_Process_Doc_For_Alignment**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Master_Process_Doc_For_Alignment(TargetDocument As Document)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||


---
VBA Procedure: **Master_Process_CtDoc_For_Alignment**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Master_Process_CtDoc_For_Alignment()*  

**no arguments required for this procedure**


---
VBA Procedure: **Master_Sort_ActiveDoc_FadoGlossary**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Master_Sort_ActiveDoc_FadoGlossary()*  

**no arguments required for this procedure**


---
VBA Procedure: **Master_Sort_TargetDoc_forAlignment_FadoGlossary**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Master_Sort_TargetDoc_forAlignment_FadoGlossary()*  

**no arguments required for this procedure**


---
VBA Procedure: **GetAll_MainIDs_ToNewDocument_Name**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub GetAll_MainIDs_ToNewDocument_Name()*  

**no arguments required for this procedure**


---
VBA Procedure: **Extract_abdCheck_All_CrossReferences**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Extract_abdCheck_All_CrossReferences()*  

**no arguments required for this procedure**


---
VBA Procedure: **Get_All_IDs_andHyperlinks**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Get_All_IDs_andHyperlinks(TargetDocument As Document) As String()*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||


---
VBA Procedure: **Extract_AndList_AllHyperlinks**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Extract_AndList_AllHyperlinks()*  

**no arguments required for this procedure**


---
VBA Procedure: **All_IDs_Bookmarks_LocationCheck_CtDoc**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub All_IDs_Bookmarks_LocationCheck_CtDoc()*  

**no arguments required for this procedure**


---
VBA Procedure: **Check_andFix_allIDs**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Check_andFix_allIDs(TargetDocument As Document) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||


---
VBA Procedure: **Repare_IDs_forBookmark**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Repare_IDs_forBookmark(TargetBookmark As Bookmark)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetBookmark|Bookmark|False||


---
VBA Procedure: **Repair_IDs_fromCollection**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Repair_IDs_fromCollection(TargetBookmark As Bookmark, TargetParagraphsCollection As Collection)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetBookmark|Bookmark|False||
TargetParagraphsCollection|Collection|False||


---
VBA Procedure: **Repair_MainIDParagraph**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Repair_MainIDParagraph(ByRef TargetParagraph As Paragraph)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByRef|Paragraph|False||


---
VBA Procedure: **Get_IDsParagraphs_forRange**  
Type: **Function**  
Returns: **Collection**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Get_IDsParagraphs_forRange(TargetRange As Range) As Collection*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetRange|Range|False||


---
VBA Procedure: **Set_Top_Bookmarks**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Set_Top_Bookmarks(TargetDocument As Document, Remove_MultiRow_Tables As Boolean)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||
Remove_MultiRow_Tables|Boolean|False||


---
VBA Procedure: **Set_MainBookmarks**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Set_MainBookmarks(TargetDocument As Document, HowManyTopTables As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||
HowManyTopTables|Integer|False||


---
VBA Procedure: **Set_MainBookmarks_accToSortingOrder**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Set_MainBookmarks_accToSortingOrder(TargetDocument As Document, OriginalDoc_SortingOrder() As String, HowManyTopTables As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||
OriginalDoc_SortingOrder|Variant|False||
HowManyTopTables|Integer|False||


---
VBA Procedure: **Get_MainIDs_Original_SortingOrder_forAllBookmarks**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Get_MainIDs_Original_SortingOrder_forAllBookmarks(TargetDocument As Document, BookmarksPrefix As String, NumberOfBookmarks As Integer) As String()*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||
BookmarksPrefix|String|False||
NumberOfBookmarks|Integer|False||


---
VBA Procedure: **Create_Fado_SortingKeys_ForAllBookmarks**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Create_Fado_SortingKeys_ForAllBookmarks(BookmarksPrefix As String, NumberOfBookmarks As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
BookmarksPrefix|String|False||
NumberOfBookmarks|Integer|False||


---
VBA Procedure: **Bkm_Sort_byName**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Bkm_Sort_byName(TargetDocument As Document, BookmarksNumber As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||
BookmarksNumber|Integer|False||


---
VBA Procedure: **Bkm_Sort**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Bkm_Sort(BookmarksPrefix As String, Optional BookmarksNumber)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
BookmarksPrefix|String|False||
BookmarksNumber|Variant|True||


---
VBA Procedure: **All_OLEShapes_Count**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub All_OLEShapes_Count()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllOleShapes_Extract**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllOleShapes_Extract()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllOle_InlineShapes_ConvertToPicture**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllOle_InlineShapes_ConvertToPicture()*  

**no arguments required for this procedure**


---
VBA Procedure: **Highlight_All_TopTables_CvTxt_Rest**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Highlight_All_TopTables_CvTxt_Rest(TargetDocument As Document)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||


---
VBA Procedure: **Count_TopTables_CtDoc**  
Type: **Function**  
Returns: **Integer**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Count_TopTables_CtDoc() As Integer*  

**no arguments required for this procedure**


---
VBA Procedure: **Green_All_GoodTopTables**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Green_All_GoodTopTables()*  

**no arguments required for this procedure**


---
VBA Procedure: **Gray_All_GoodTopTables**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Gray_All_GoodTopTables()*  

**no arguments required for this procedure**


---
VBA Procedure: **Identify_First_MultiRow_TopTable**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Identify_First_MultiRow_TopTable()*  

**no arguments required for this procedure**


---
VBA Procedure: **Count_TopTables**  
Type: **Function**  
Returns: **Integer**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Count_TopTables(TargetDocument As Document) As Integer*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||


---
VBA Procedure: **Select_BadColorIndex_TopTable**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Select_BadColorIndex_TopTable()*  

**no arguments required for this procedure**


---
VBA Procedure: **Remove_AllTopTables**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Remove_AllTopTables()*  

**no arguments required for this procedure**


---
VBA Procedure: **Empty_NonInlineShape_HyperLinks_Delete**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Empty_NonInlineShape_HyperLinks_Delete()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllTables_FitContents**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllTables_FitContents()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllReferences_InMainText_ToRefStyle**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllReferences_InMainText_ToRefStyle()*  

**no arguments required for this procedure**


---
VBA Procedure: **CoverNote_Remove**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub CoverNote_Remove()*  

**no arguments required for this procedure**


---
VBA Procedure: **Empty_OneRow_Tables_Remove**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Empty_OneRow_Tables_Remove()*  

**no arguments required for this procedure**


---
VBA Procedure: **Empty_MultiRowTables_Count**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Empty_MultiRowTables_Count()*  

**no arguments required for this procedure**


---
VBA Procedure: **Empty_MultiRowTables_Remove**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Empty_MultiRowTables_Remove()*  

**no arguments required for this procedure**


---
VBA Procedure: **IsTable_Empty**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function IsTable_Empty(TargetTable As Table) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetTable|Table|False||


---
VBA Procedure: **AllVBCr_Pairs_ToVbCr**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllVBCr_Pairs_ToVbCr(TargetDocument As Document)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||


---
VBA Procedure: **AllMultiple_VBCrs_To_Single**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllMultiple_VBCrs_To_Single(TargetDocument As Document)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||


---
VBA Procedure: **Get_First_NonEmpty_NonNumeric_Paragraph**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Get_First_NonEmpty_NonNumeric_Paragraph(TargetDocument As Document, TargetBookmark As Bookmark) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||
TargetBookmark|Bookmark|False||


---
VBA Procedure: **ExtractAll_NonEmpty_NonNumeric_FirstParagraphs_ToNewDoc**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub ExtractAll_NonEmpty_NonNumeric_FirstParagraphs_ToNewDoc()*  

**no arguments required for this procedure**


---
VBA Procedure: **Get_MainID_fromBookmark**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Get_MainID_fromBookmark(TargetDocument As Document, TargetBookmark As Bookmark) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||
TargetBookmark|Bookmark|False||


---
VBA Procedure: **All_MainIDs_ToHeading1_Style**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub All_MainIDs_ToHeading1_Style()*  

**no arguments required for this procedure**


---
VBA Procedure: **CurrentSelection_CrossRef_Add**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub CurrentSelection_CrossRef_Add()*  

**no arguments required for this procedure**


---
VBA Procedure: **Replace_WrongSubAdress_to000_inAllHyperLinks**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Replace_WrongSubAdress_to000_inAllHyperLinks(CurrentWrongAdress As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
CurrentWrongAdress|String|False||


---
VBA Procedure: **Remove_1_from_AllHyperlinks**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Remove_1_from_AllHyperlinks()*  

**no arguments required for this procedure**


---
VBA Procedure: **All_Bookmarks_WithPrefix_Remove**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub All_Bookmarks_WithPrefix_Remove(Prefix As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Prefix|String|False||


---
VBA Procedure: **All_Bookmarks_Remove**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub All_Bookmarks_Remove(TargetDocument As Document)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||


---
VBA Procedure: **RemoveAll_MultiRows_Tables**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub RemoveAll_MultiRows_Tables()*  

**no arguments required for this procedure**


---
VBA Procedure: **Count_MultipleRows_Tables**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Count_MultipleRows_Tables()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllHeadings1_IDS_Highlight**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllHeadings1_IDS_Highlight()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllHeadings1_IDS_GreyArial14Bold**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllHeadings1_IDS_GreyArial14Bold()*  

**no arguments required for this procedure**


---
VBA Procedure: **All_Heading2_IDS_GreyArial14NoBold**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub All_Heading2_IDS_GreyArial14NoBold()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllHeading1_IDS_Highlight**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllHeading1_IDS_Highlight()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllHeading1_IDS_MakeBookmark_Before**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllHeading1_IDS_MakeBookmark_Before()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllHeading2_IDS_Highlight**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllHeading2_IDS_Highlight()*  

**no arguments required for this procedure**


---
VBA Procedure: **Extract_AllHeadings1_Titles_ToNewDoc**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Extract_AllHeadings1_Titles_ToNewDoc()*  

**no arguments required for this procedure**


---
VBA Procedure: **Contents_Check_AndPutTogether**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Contents_Check_AndPutTogether()*  

**no arguments required for this procedure**


---
VBA Procedure: **Swap_IDS_With_Titles**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Swap_IDS_With_Titles()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllHeadings1_Highlight**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllHeadings1_Highlight()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllHeadings2_Highlight**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllHeadings2_Highlight()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllHeadings3_Highlight**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllHeadings3_Highlight()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllParagraphs_WithBullet_ToHeading3**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllParagraphs_WithBullet_ToHeading3()*  

**no arguments required for this procedure**


---
VBA Procedure: **All_Heading2_ToHeading1_Style_Change**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub All_Heading2_ToHeading1_Style_Change()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllHeadings1_MarkIf_BeforeLessThen4Enters**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllHeadings1_MarkIf_BeforeLessThen4Enters()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllTables_Remove**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllTables_Remove()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllTables_ConvertToText**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllTables_ConvertToText(TargetDocument As Document)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||


---
VBA Procedure: **AllTables_ConvertToText_CtDoc**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllTables_ConvertToText_CtDoc()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllInlineImagesDelete**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllInlineImagesDelete(TargetDocument As Document)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||


---
VBA Procedure: **AllShapesDelete**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllShapesDelete()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllPicturesDelete**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllPicturesDelete()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllPageBreaks_ToEnters**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllPageBreaks_ToEnters(TargetDocument As Document)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||


---
VBA Procedure: **All_Heading1_OneEmptyParagraph_Before**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub All_Heading1_OneEmptyParagraph_Before()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllHeadings1_BetweenENDoubleQuotes**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllHeadings1_BetweenENDoubleQuotes()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllEmptyParagraphs_ToNormalStyle**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllEmptyParagraphs_ToNormalStyle()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllHeading1_KeepWithNext**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllHeading1_KeepWithNext()*  

**no arguments required for this procedure**


---
VBA Procedure: **Show_DoubleQuotesIn_Headings1**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Show_DoubleQuotesIn_Headings1()*  

**no arguments required for this procedure**


---
VBA Procedure: **AllHeadings1_NoQuotes**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AllHeadings1_NoQuotes()*  

**no arguments required for this procedure**


---
VBA Procedure: **Temporarily_DoubleQuotes_ToFrenchQuotes_InNonHeadings1**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Temporarily_DoubleQuotes_ToFrenchQuotes_InNonHeadings1()*  

**no arguments required for this procedure**


---
VBA Procedure: **all20Sized_ToHeadings1**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub all20Sized_ToHeadings1()*  

**no arguments required for this procedure**


---
VBA Procedure: **countHeadings1**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub countHeadings1()*  

**no arguments required for this procedure**


---
VBA Procedure: **RemoveAll_PagebreakBefores**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub RemoveAll_PagebreakBefores(TargetDocument As Document)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||


---
VBA Procedure: **Get_All_Hidden_IDBookmarks**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Get_All_Hidden_IDBookmarks(TargetDocument As Document) As String()*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||


---
VBA Procedure: **Get_IDs_Bookmarks_LocationCheck**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Get_IDs_Bookmarks_LocationCheck(TargetDocument As Document) As String()*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetDocument|Document|False||
