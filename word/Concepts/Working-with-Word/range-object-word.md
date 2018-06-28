---
title: Range Object (Word)
keywords: vbawd10.chm2398
f1_keywords:
- vbawd10.chm2398
ms.prod: word
api_name:
- Word.Range
ms.assetid: 15a7a1c4-5f3f-5b6e-60e9-29688de3f274
ms.date: 06/08/2017
---


# Range Object (Word)

Represents a contiguous area in a document. Each  **Range** object is defined by a starting and ending character position.


## Remarks

Similar to the way bookmarks are used in a document,  **Range** objects are used in Visual Basic procedures to identify specific portions of a document. However, unlike a bookmark, a **Range** object only exists while the procedure that defined it is running. **Range** objects are independent of the selection. That is, you can define and manipulate a range without changing the selection. You can also define multiple ranges in a document, while there can be only one selection per pane.

Use the  **Range** method to return a **Range** object defined by the given starting and ending character positions. The following example returns a **Range** object that refers to the first 10 characters in the active document.




```
Set myRange = ActiveDocument.Range(Start:=0, End:=10)
```

Use the  **Range** property to return a **Range** object defined by the beginning and end of another object. The **Range** property applies to many objects (for example, **Paragraph**, **Bookmark**, and **Cell** ). The following example returns a **Range** object that refers to the first paragraph in the active document.




```
Set aRange = ActiveDocument.Paragraphs(1).Range
```

The following example returns a  **Range** object that refers to the second through fourth paragraphs in the active document




```
Set aRange = ActiveDocument.Range( _ 
 Start:=ActiveDocument.Paragraphs(2).Range.Start, _ 
 End:=ActiveDocument.Paragraphs(4).Range.End)
```

For more information about working with  **Range** objects, see [Working with Range Objects](../Miscellaneous/working-with-range-objects.md).


## Methods



|**Name**|
|:-----|
|[AutoFormat](../../../api/Word.Range.AutoFormat.md)|
|[Calculate](../../../api/Word.Range.Calculate.md)|
|[CheckGrammar](../../../api/Word.Range.CheckGrammar.md)|
|[CheckSpelling](../../../api/Word.Range.CheckSpelling.md)|
|[CheckSynonyms](../../../api/Word.Range.CheckSynonyms.md)|
|[Collapse](../../../api/Word.Range.Collapse.md)|
|[ComputeStatistics](../../../api/Word.Range.ComputeStatistics.md)|
|[ConvertHangulAndHanja](../../../api/Word.Range.ConvertHangulAndHanja.md)|
|[ConvertToTable](../../../api/Word.Range.ConvertToTable.md)|
|[Copy](../../../api/Word.Range.Copy.md)|
|[CopyAsPicture](../../../api/Word.Range.CopyAsPicture.md)|
|[Cut](../../../api/Word.Range.Cut.md)|
|[Delete](../../../api/Word.Range.Delete.md)|
|[DetectLanguage](../../../api/Word.Range.DetectLanguage.md)|
|[EndOf](../../../api/Word.Range.EndOf.md)|
|[Expand](../../../api/Word.Range.Expand.md)|
|[ExportAsFixedFormat](../../../api/Word.Range.ExportAsFixedFormat.md)|
|[ExportFragment](../../../api/Word.Range.ExportFragment.md)|
|[GetSpellingSuggestions](../../../api/Word.Range.GetSpellingSuggestions.md)|
|[GoTo](../../../api/Word.Range.GoTo.md)|
|[GoToEditableRange](../../../api/Word.Range.GoToEditableRange.md)|
|[GoToNext](../../../api/Word.Range.GoToNext.md)|
|[GoToPrevious](../../../api/Word.Range.GoToPrevious.md)|
|[ImportFragment](../../../api/Word.Range.ImportFragment.md)|
|[InRange](../../../api/Word.Range.InRange.md)|
|[InsertAfter](../../../api/Word.Range.InsertAfter.md)|
|[InsertAlignmentTab](../../../api/Word.Range.InsertAlignmentTab.md)|
|[InsertAutoText](../../../api/Word.Range.InsertAutoText.md)|
|[InsertBefore](../../../api/Word.Range.InsertBefore.md)|
|[InsertBreak](../../../api/Word.Range.InsertBreak.md)|
|[InsertCaption](../../../api/Word.Range.InsertCaption.md)|
|[InsertCrossReference](../../../api/Word.Range.InsertCrossReference.md)|
|[InsertDatabase](../../../api/Word.Range.InsertDatabase.md)|
|[InsertDateTime](../../../api/Word.Range.InsertDateTime.md)|
|[InsertFile](../../../api/Word.Range.InsertFile.md)|
|[InsertParagraph](../../../api/Word.Range.InsertParagraph.md)|
|[InsertParagraphAfter](../../../api/Word.Range.InsertParagraphAfter.md)|
|[InsertParagraphBefore](../../../api/Word.Range.InsertParagraphBefore.md)|
|[InsertSymbol](../../../api/Word.Range.InsertSymbol.md)|
|[InsertXML](../../../api/Word.Range.InsertXML.md)|
|[InStory](../../../api/Word.Range.InStory.md)|
|[IsEqual](../../../api/Word.Range.IsEqual.md)|
|[LookupNameProperties](../../../api/Word.Range.LookupNameProperties.md)|
|[ModifyEnclosure](../../../api/Word.Range.ModifyEnclosure.md)|
|[Move](../../../api/Word.Range.Move.md)|
|[MoveEnd](../../../api/Word.Range.MoveEnd.md)|
|[MoveEndUntil](../../../api/Word.Range.MoveEndUntil.md)|
|[MoveEndWhile](../../../api/Word.Range.MoveEndWhile.md)|
|[MoveStart](../../../api/Word.Range.MoveStart.md)|
|[MoveStartUntil](../../../api/Word.Range.MoveStartUntil.md)|
|[MoveStartWhile](../../../api/Word.Range.MoveStartWhile.md)|
|[MoveUntil](../../../api/Word.Range.MoveUntil.md)|
|[MoveWhile](../../../api/Word.Range.MoveWhile.md)|
|[Next](../../../api/Word.Range.Next.md)|
|[NextSubdocument](../../../api/Word.Range.NextSubdocument.md)|
|[Paste](../../../api/Word.Range.Paste.md)|
|[PasteAndFormat](../../../api/Word.Range.PasteAndFormat.md)|
|[PasteAppendTable](../../../api/Word.Range.PasteAppendTable.md)|
|[PasteAsNestedTable](../../../api/Word.Range.PasteAsNestedTable.md)|
|[PasteExcelTable](../../../api/Word.Range.PasteExcelTable.md)|
|[PasteSpecial](../../../api/Word.Range.PasteSpecial.md)|
|[PhoneticGuide](../../../api/Word.Range.PhoneticGuide.md)|
|[Previous](../../../api/Word.Range.Previous.md)|
|[PreviousSubdocument](../../../api/Word.Range.PreviousSubdocument.md)|
|[Relocate](../../../api/Word.Range.Relocate.md)|
|[Select](../../../api/Word.Range.Select.md)|
|[SetListLevel](../../../api/Word.Range.SetListLevel.md)|
|[SetRange](../../../api/Word.Range.SetRange.md)|
|[Sort](../../../api/Word.Range.Sort.md)|
|[SortAscending](../../../api/Word.Range.SortAscending.md)|
|[SortByHeadings](../../../api/Word.range.sortbyheadings.md)|
|[SortDescending](../../../api/Word.Range.SortDescending.md)|
|[StartOf](../../../api/Word.Range.StartOf.md)|
|[TCSCConverter](../../../api/Word.Range.TCSCConverter.md)|
|[WholeStory](../../../api/Word.Range.WholeStory.md)|

## Properties



|**Name**|
|:-----|
|[Application](../../../api/Word.Range.Application.md)|
|[Bold](../../../api/Word.Range.Bold.md)|
|[BoldBi](../../../api/Word.Range.BoldBi.md)|
|[BookmarkID](../../../api/Word.Range.BookmarkID.md)|
|[Bookmarks](../../../api/Word.Range.Bookmarks.md)|
|[Borders](../../../api/Word.Range.Borders.md)|
|[Case](../../../api/Word.Range.Case.md)|
|[Cells](../../../api/Word.Range.Cells.md)|
|[Characters](../../../api/Word.Range.Characters.md)|
|[CharacterStyle](../../../api/Word.Range.CharacterStyle.md)|
|[CharacterWidth](../../../api/Word.Range.CharacterWidth.md)|
|[Columns](../../../api/Word.Range.Columns.md)|
|[CombineCharacters](../../../api/Word.Range.CombineCharacters.md)|
|[Comments](../../../api/Word.Range.Comments.md)|
|[Conflicts](../../../api/Word.Range.Conflicts.md)|
|[ContentControls](../../../api/Word.Range.ContentControls.md)|
|[Creator](../../../api/Word.Range.Creator.md)|
|[DisableCharacterSpaceGrid](../../../api/Word.Range.DisableCharacterSpaceGrid.md)|
|[Document](../../../api/Word.Range.Document.md)|
|[Duplicate](../../../api/Word.Range.Duplicate.md)|
|[Editors](../../../api/Word.Range.Editors.md)|
|[EmphasisMark](../../../api/Word.Range.EmphasisMark.md)|
|[End](../../../api/Word.Range.End.md)|
|[EndnoteOptions](../../../api/Word.Range.EndnoteOptions.md)|
|[Endnotes](../../../api/Word.Range.Endnotes.md)|
|[EnhMetaFileBits](../../../api/Word.Range.EnhMetaFileBits.md)|
|[Fields](../../../api/Word.Range.Fields.md)|
|[Find](../../../api/Word.Range.Find.md)|
|[FitTextWidth](../../../api/Word.Range.FitTextWidth.md)|
|[Font](../../../api/Word.Range.Font.md)|
|[FootnoteOptions](../../../api/Word.Range.FootnoteOptions.md)|
|[Footnotes](../../../api/Word.Range.Footnotes.md)|
|[FormattedText](../../../api/Word.Range.FormattedText.md)|
|[FormFields](../../../api/Word.Range.FormFields.md)|
|[Frames](../../../api/Word.Range.Frames.md)|
|[GrammarChecked](../../../api/Word.Range.GrammarChecked.md)|
|[GrammaticalErrors](../../../api/Word.Range.GrammaticalErrors.md)|
|[HighlightColorIndex](../../../api/Word.Range.HighlightColorIndex.md)|
|[HorizontalInVertical](../../../api/Word.Range.HorizontalInVertical.md)|
|[HTMLDivisions](../../../api/Word.Range.HTMLDivisions.md)|
|[Hyperlinks](../../../api/Word.Range.Hyperlinks.md)|
|[ID](../../../api/Word.Range.ID.md)|
|[Information](../../../api/Word.Range.Information.md)|
|[InlineShapes](../../../api/Word.Range.InlineShapes.md)|
|[IsEndOfRowMark](../../../api/Word.Range.IsEndOfRowMark.md)|
|[Italic](../../../api/Word.Range.Italic.md)|
|[ItalicBi](../../../api/Word.Range.ItalicBi.md)|
|[Kana](../../../api/Word.Range.Kana.md)|
|[LanguageDetected](../../../api/Word.Range.LanguageDetected.md)|
|[LanguageID](../../../api/Word.Range.LanguageID.md)|
|[LanguageIDFarEast](../../../api/Word.Range.LanguageIDFarEast.md)|
|[LanguageIDOther](../../../api/Word.Range.LanguageIDOther.md)|
|[ListFormat](../../../api/Word.Range.ListFormat.md)|
|[ListParagraphs](../../../api/Word.Range.ListParagraphs.md)|
|[ListStyle](../../../api/Word.Range.ListStyle.md)|
|[Locks](../../../api/Word.Range.Locks.md)|
|[NextStoryRange](../../../api/Word.Range.NextStoryRange.md)|
|[NoProofing](../../../api/Word.Range.NoProofing.md)|
|[OMaths](../../../api/Word.Range.OMaths.md)|
|[Orientation](../../../api/Word.Range.Orientation.md)|
|[PageSetup](../../../api/Word.Range.PageSetup.md)|
|[ParagraphFormat](../../../api/Word.Range.ParagraphFormat.md)|
|[Paragraphs](../../../api/Word.Range.Paragraphs.md)|
|[ParagraphStyle](../../../api/Word.Range.ParagraphStyle.md)|
|[Parent](../../../api/Word.Range.Parent.md)|
|[ParentContentControl](../../../api/Word.Range.ParentContentControl.md)|
|[PreviousBookmarkID](../../../api/Word.Range.PreviousBookmarkID.md)|
|[ReadabilityStatistics](../../../api/Word.Range.ReadabilityStatistics.md)|
|[Revisions](../../../api/Word.Range.Revisions.md)|
|[Rows](../../../api/Word.Range.Rows.md)|
|[Scripts](../../../api/Word.Range.Scripts.md)|
|[Sections](../../../api/Word.Range.Sections.md)|
|[Sentences](../../../api/Word.Range.Sentences.md)|
|[Shading](../../../api/Word.Range.Shading.md)|
|[ShapeRange](../../../api/Word.Range.ShapeRange.md)|
|[ShowAll](../../../api/Word.Range.ShowAll.md)|
|[SpellingChecked](../../../api/Word.Range.SpellingChecked.md)|
|[SpellingErrors](../../../api/Word.Range.SpellingErrors.md)|
|[Start](../../../api/Word.Range.Start.md)|
|[StoryLength](../../../api/Word.Range.StoryLength.md)|
|[StoryType](../../../api/Word.Range.StoryType.md)|
|[Style](../../../api/Word.Range.Style.md)|
|[Subdocuments](../../../api/Word.Range.Subdocuments.md)|
|[SynonymInfo](../../../api/Word.Range.SynonymInfo.md)|
|[Tables](../../../api/Word.Range.Tables.md)|
|[TableStyle](../../../api/Word.Range.TableStyle.md)|
|[Text](../../../api/Word.Range.Text.md)|
|[TextRetrievalMode](../../../api/Word.Range.TextRetrievalMode.md)|
|[TextVisibleOnScreen](../../../api/Word.range.textvisibleonscreen.md)|
|[TopLevelTables](../../../api/Word.Range.TopLevelTables.md)|
|[TwoLinesInOne](../../../api/Word.Range.TwoLinesInOne.md)|
|[Underline](../../../api/Word.Range.Underline.md)|
|[Updates](../../../api/Word.Range.Updates.md)|
|[WordOpenXML](../../../api/Word.Range.WordOpenXML.md)|
|[Words](../../../api/Word.Range.Words.md)|
|[XML](../../../api/Word.Range.XML.md)|

## See also


#### Other resources


[Word Object Model Reference](../../../api/overview/object-model-word-vba-reference.md)

