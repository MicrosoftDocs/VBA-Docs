---
title: Range object (Word)
keywords: vbawd10.chm2398
f1_keywords:
- vbawd10.chm2398
ms.prod: word
api_name:
- Word.Range
ms.assetid: 15a7a1c4-5f3f-5b6e-60e9-29688de3f274
ms.date: 05/23/2019
localization_priority: Normal
---


# Range object (Word)

Represents a contiguous area in a document. Each **Range** object is defined by a starting and ending character position.


## Remarks

Similar to the way bookmarks are used in a document, **Range** objects are used in Visual Basic procedures to identify specific portions of a document. However, unlike a bookmark, a **Range** object only exists while the procedure that defined it is running. **Range** objects are independent of the selection. That is, you can define and manipulate a range without changing the selection. You can also define multiple ranges in a document, while there can be only one selection per pane.

Use the **Range** method to return a **Range** object defined by the given starting and ending character positions. The following example returns a **Range** object that refers to the first 10 characters in the active document.

```vb
Set myRange = ActiveDocument.Range(Start:=0, End:=10)
```

Use the **Range** property to return a **Range** object defined by the beginning and end of another object. The **Range** property applies to many objects (for example, **Paragraph**, **Bookmark**, and **Cell**). The following example returns a **Range** object that refers to the first paragraph in the active document.

```vb
Set aRange = ActiveDocument.Paragraphs(1).Range
```

The following example returns a **Range** object that refers to the second through fourth paragraphs in the active document.

```vb
Set aRange = ActiveDocument.Range( _ 
 Start:=ActiveDocument.Paragraphs(2).Range.Start, _ 
 End:=ActiveDocument.Paragraphs(4).Range.End)
```

For more information about working with **Range** objects, see [Working with Range objects](../word/Concepts/Working-with-Word/working-with-range-objects.md).


## Methods

- [AutoFormat](Word.Range.AutoFormat.md)
- [Calculate](Word.Range.Calculate.md)
- [CheckGrammar](Word.Range.CheckGrammar.md)
- [CheckSpelling](Word.Range.CheckSpelling.md)
- [CheckSynonyms](Word.Range.CheckSynonyms.md)
- [Collapse](Word.Range.Collapse.md)
- [ComputeStatistics](Word.Range.ComputeStatistics.md)
- [ConvertHangulAndHanja](Word.Range.ConvertHangulAndHanja.md)
- [ConvertToTable](Word.Range.ConvertToTable.md)
- [Copy](Word.Range.Copy.md)
- [CopyAsPicture](Word.Range.CopyAsPicture.md)
- [Cut](Word.Range.Cut.md)
- [Delete](Word.Range.Delete.md)
- [DetectLanguage](Word.Range.DetectLanguage.md)
- [EndOf](Word.Range.EndOf.md)
- [Expand](Word.Range.Expand.md)
- [ExportAsFixedFormat](Word.Range.ExportAsFixedFormat.md)
- [ExportAsFixedFormat2](Word.Range.ExportAsFixedFormat2.md)
- [ExportFragment](Word.Range.ExportFragment.md)
- [GetSpellingSuggestions](Word.Range.GetSpellingSuggestions.md)
- [GoTo](Word.Range.GoTo.md)
- [GoToEditableRange](Word.Range.GoToEditableRange.md)
- [GoToNext](Word.Range.GoToNext.md)
- [GoToPrevious](Word.Range.GoToPrevious.md)
- [ImportFragment](Word.Range.ImportFragment.md)
- [InRange](Word.Range.InRange.md)
- [InsertAfter](Word.Range.InsertAfter.md)
- [InsertAlignmentTab](Word.Range.InsertAlignmentTab.md)
- [InsertAutoText](Word.Range.InsertAutoText.md)
- [InsertBefore](Word.Range.InsertBefore.md)
- [InsertBreak](Word.Range.InsertBreak.md)
- [InsertCaption](Word.Range.InsertCaption.md)
- [InsertCrossReference](Word.Range.InsertCrossReference.md)
- [InsertDatabase](Word.Range.InsertDatabase.md)
- [InsertDateTime](Word.Range.InsertDateTime.md)
- [InsertFile](Word.Range.InsertFile.md)
- [InsertParagraph](Word.Range.InsertParagraph.md)
- [InsertParagraphAfter](Word.Range.InsertParagraphAfter.md)
- [InsertParagraphBefore](Word.Range.InsertParagraphBefore.md)
- [InsertSymbol](Word.Range.InsertSymbol.md)
- [InsertXML](Word.Range.InsertXML.md)
- [InStory](Word.Range.InStory.md)
- [IsEqual](Word.Range.IsEqual.md)
- [LookupNameProperties](Word.Range.LookupNameProperties.md)
- [ModifyEnclosure](Word.Range.ModifyEnclosure.md)
- [Move](Word.Range.Move.md)
- [MoveEnd](Word.Range.MoveEnd.md)
- [MoveEndUntil](Word.Range.MoveEndUntil.md)
- [MoveEndWhile](Word.Range.MoveEndWhile.md)
- [MoveStart](Word.Range.MoveStart.md)
- [MoveStartUntil](Word.Range.MoveStartUntil.md)
- [MoveStartWhile](Word.Range.MoveStartWhile.md)
- [MoveUntil](Word.Range.MoveUntil.md)
- [MoveWhile](Word.Range.MoveWhile.md)
- [Next](Word.Range.Next.md)
- [NextSubdocument](Word.Range.NextSubdocument.md)
- [Paste](Word.Range.Paste.md)
- [PasteAndFormat](Word.Range.PasteAndFormat.md)
- [PasteAppendTable](Word.Range.PasteAppendTable.md)
- [PasteAsNestedTable](Word.Range.PasteAsNestedTable.md)
- [PasteExcelTable](Word.Range.PasteExcelTable.md)
- [PasteSpecial](Word.Range.PasteSpecial.md)
- [PhoneticGuide](Word.Range.PhoneticGuide.md)
- [Previous](Word.Range.Previous.md)
- [PreviousSubdocument](Word.Range.PreviousSubdocument.md)
- [Relocate](Word.Range.Relocate.md)
- [Select](Word.Range.Select.md)
- [SetListLevel](Word.Range.SetListLevel.md)
- [SetRange](Word.Range.SetRange.md)
- [Sort](Word.Range.Sort.md)
- [SortAscending](Word.Range.SortAscending.md)
- [SortByHeadings](Word.range.sortbyheadings.md)
- [SortDescending](Word.Range.SortDescending.md)
- [StartOf](Word.Range.StartOf.md)
- [TCSCConverter](Word.Range.TCSCConverter.md)
- [WholeStory](Word.Range.WholeStory.md)

## Properties

- [Application](Word.Range.Application.md)
- [Bold](Word.Range.Bold.md)
- [BoldBi](Word.Range.BoldBi.md)
- [BookmarkID](Word.Range.BookmarkID.md)
- [Bookmarks](Word.Range.Bookmarks.md)
- [Borders](Word.Range.Borders.md)
- [Case](Word.Range.Case.md)
- [Cells](Word.Range.Cells.md)
- [Characters](Word.Range.Characters.md)
- [CharacterStyle](Word.Range.CharacterStyle.md)
- [CharacterWidth](Word.Range.CharacterWidth.md)
- [Columns](Word.Range.Columns.md)
- [CombineCharacters](Word.Range.CombineCharacters.md)
- [Comments](Word.Range.Comments.md)
- [Conflicts](Word.Range.Conflicts.md)
- [ContentControls](Word.Range.ContentControls.md)
- [Creator](Word.Range.Creator.md)
- [DisableCharacterSpaceGrid](Word.Range.DisableCharacterSpaceGrid.md)
- [Document](Word.Range.Document.md)
- [Duplicate](Word.Range.Duplicate.md)
- [Editors](Word.Range.Editors.md)
- [EmphasisMark](Word.Range.EmphasisMark.md)
- [End](Word.Range.End.md)
- [EndnoteOptions](Word.Range.EndnoteOptions.md)
- [Endnotes](Word.Range.Endnotes.md)
- [EnhMetaFileBits](Word.Range.EnhMetaFileBits.md)
- [Fields](Word.Range.Fields.md)
- [Find](Word.Range.Find.md)
- [FitTextWidth](Word.Range.FitTextWidth.md)
- [Font](Word.Range.Font.md)
- [FootnoteOptions](Word.Range.FootnoteOptions.md)
- [Footnotes](Word.Range.Footnotes.md)
- [FormattedText](Word.Range.FormattedText.md)
- [FormFields](Word.Range.FormFields.md)
- [Frames](Word.Range.Frames.md)
- [GrammarChecked](Word.Range.GrammarChecked.md)
- [GrammaticalErrors](Word.Range.GrammaticalErrors.md)
- [HighlightColorIndex](Word.Range.HighlightColorIndex.md)
- [HorizontalInVertical](Word.Range.HorizontalInVertical.md)
- [HTMLDivisions](Word.Range.HTMLDivisions.md)
- [Hyperlinks](Word.Range.Hyperlinks.md)
- [ID](Word.Range.ID.md)
- [Information](Word.Range.Information.md)
- [InlineShapes](Word.Range.InlineShapes.md)
- [IsEndOfRowMark](Word.Range.IsEndOfRowMark.md)
- [Italic](Word.Range.Italic.md)
- [ItalicBi](Word.Range.ItalicBi.md)
- [Kana](Word.Range.Kana.md)
- [LanguageDetected](Word.Range.LanguageDetected.md)
- [LanguageID](Word.Range.LanguageID.md)
- [LanguageIDFarEast](Word.Range.LanguageIDFarEast.md)
- [LanguageIDOther](Word.Range.LanguageIDOther.md)
- [ListFormat](Word.Range.ListFormat.md)
- [ListParagraphs](Word.Range.ListParagraphs.md)
- [ListStyle](Word.Range.ListStyle.md)
- [Locks](Word.Range.Locks.md)
- [NextStoryRange](Word.Range.NextStoryRange.md)
- [NoProofing](Word.Range.NoProofing.md)
- [OMaths](Word.Range.OMaths.md)
- [Orientation](Word.Range.Orientation.md)
- [PageSetup](Word.Range.PageSetup.md)
- [ParagraphFormat](Word.Range.ParagraphFormat.md)
- [Paragraphs](Word.Range.Paragraphs.md)
- [ParagraphStyle](Word.Range.ParagraphStyle.md)
- [Parent](Word.Range.Parent.md)
- [ParentContentControl](Word.Range.ParentContentControl.md)
- [PreviousBookmarkID](Word.Range.PreviousBookmarkID.md)
- [ReadabilityStatistics](Word.Range.ReadabilityStatistics.md)
- [Revisions](Word.Range.Revisions.md)
- [Rows](Word.Range.Rows.md)
- [Scripts](Word.Range.Scripts.md)
- [Sections](Word.Range.Sections.md)
- [Sentences](Word.Range.Sentences.md)
- [Shading](Word.Range.Shading.md)
- [ShapeRange](Word.Range.ShapeRange.md)
- [ShowAll](Word.Range.ShowAll.md)
- [SpellingChecked](Word.Range.SpellingChecked.md)
- [SpellingErrors](Word.Range.SpellingErrors.md)
- [Start](Word.Range.Start.md)
- [StoryLength](Word.Range.StoryLength.md)
- [StoryType](Word.Range.StoryType.md)
- [Style](Word.Range.Style.md)
- [Subdocuments](Word.Range.Subdocuments.md)
- [SynonymInfo](Word.Range.SynonymInfo.md)
- [Tables](Word.Range.Tables.md)
- [TableStyle](Word.Range.TableStyle.md)
- [Text](Word.Range.Text.md)
- [TextRetrievalMode](Word.Range.TextRetrievalMode.md)
- [TextVisibleOnScreen](Word.range.textvisibleonscreen.md)
- [TopLevelTables](Word.Range.TopLevelTables.md)
- [TwoLinesInOne](Word.Range.TwoLinesInOne.md)
- [Underline](Word.Range.Underline.md)
- [Updates](Word.Range.Updates.md)
- [WordOpenXML](Word.Range.WordOpenXML.md)
- [Words](Word.Range.Words.md)
- [XML](Word.Range.XML.md)

## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
