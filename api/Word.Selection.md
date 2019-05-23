---
title: Selection object (Word)
keywords: vbawd10.chm2421
f1_keywords:
- vbawd10.chm2421
ms.prod: word
api_name:
- Word.Selection
ms.assetid: 7b574a91-c33e-ecfd-6783-6b7528b2ed8f
ms.date: 05/23/2019
localization_priority: Priority
---


# Selection object (Word)

Represents the current selection in a window or pane. A selection represents either a selected (or highlighted) area in the document, or it represents the insertion point if nothing in the document is selected. There can be only one **Selection** object per document window pane, and only one **Selection** object in the entire application can be active.


## Remarks

Use the **Selection** property to return the **Selection** object. If no object qualifier is used with the **Selection** property, Microsoft Word returns the selection from the active pane of the active document window. The following example copies the current selection from the active document.

```vb
Selection.Copy
```

The following example deletes the selection from the third document in the **Documents** collection. The document does not have to be active to access its current selection.

```vb
Documents(3).ActiveWindow.Selection.Cut
```

The following example copies the selection from the first pane of the active document and pastes it into the second pane.

```vb
ActiveDocument.ActiveWindow.Panes(1).Selection.Copy 
ActiveDocument.ActiveWindow.Panes(2).Selection.Paste
```

The **Text** property is the default property of the **Selection** object. Use this property to set or return the text in the current selection. The following example assigns the text in the current selection to the variable `strTemp`, removing the last character if it is a paragraph mark.

```vb
Dim strTemp as String 
 
strTemp = Selection.Text 
If Right(strTemp, 1) = vbCr Then _ 
 strTemp = Left(strTemp, Len(strTemp) - 1)
```

The **Selection** object has various methods and properties with which you can collapse, expand, or otherwise change the current selection. The following example moves the insertion point to the end of the document and selects the last three lines.

```vb
Selection.EndOf Unit:=wdStory, Extend:=wdMove 
Selection.HomeKey Unit:=wdLine, Extend:=wdExtend 
Selection.MoveUp Unit:=wdLine, Count:=2, Extend:=wdExtend
```

The **Selection** object has various methods and properties with which you can edit selected text in a document. The following example selects the first sentence in the active document and replaces it with a new paragraph.

```vb
Options.ReplaceSelection = True 
ActiveDocument.Sentences(1).Select 
Selection.TypeText "Material below is confidential." 
Selection.TypeParagraph
```

The following example deletes the last paragraph of the first document in the **Documents** collection and pastes it at the beginning of the second document.

```vb
With Documents(1) 
 .Paragraphs.Last.Range.Select 
 .ActiveWindow.Selection.Cut 
End With 
 
With Documents(2).ActiveWindow.Selection 
 .StartOf Unit:=wdStory, Extend:=wdMove 
 .Paste 
End With
```

The **Selection** object has various methods and properties with which you can change the formatting of the current selection. The following example changes the font of the current selection from Times New Roman to Tahoma.

```vb
If Selection.Font.Name = "Times New Roman" Then _ 
 Selection.Font.Name = "Tahoma"
```

Use properties like **Flags**, **Information**, and **Type** to return information about the current selection. You can use the following example in a procedure to determine whether there is anything selected in the active document; if there is not, the rest of the procedure is skipped.

```vb
If Selection.Type = wdSelectionIP Then 
 MsgBox Prompt:="You have not selected any text! Exiting procedure..." 
 Exit Sub 
End If
```

Even when a selection is collapsed to an insertion point, it is not necessarily empty. For example, the **Text** property will still return the character to the right of the insertion point; this character also appears in the **Characters** collection of the **Selection** object. However, calling methods like **Cut** or **Copy** from a collapsed selection causes an error.

It is possible for the user to select a region in a document that does not represent contiguous text (for example, when using the Alt key with the mouse). Because the behavior of such a selection can be unpredictable, you may want to include a step in your code that checks the **Type** property of a selection before performing any operations on it (`Selection.Type = wdSelectionBlock`). 

Similarly, selections that include table cells can also lead to unpredictable behavior. The **Information** property will tell you if a selection is inside a table (`Selection.Information(wdWithinTable) = True`). The following example determines if a selection is normal (for example, it is not a row or column in a table, it is not a vertical block of text); you could use it to test the current selection before performing any operations on it.

```vb
If Selection.Type <> wdSelectionNormal Then 
 MsgBox Prompt:="Not a valid selection! Exiting procedure..." 
 Exit Sub 
End If
```

Because **Range** objects share many of the same methods and properties as **Selection** objects, using **Range** objects is preferable for manipulating a document when there is not a reason to physically change the current selection. For more information about **Selection** and **Range** objects, see [Working with the Selection object](../word/Concepts/Working-with-Word/working-with-the-selection-object.md) and [Working with Range objects](../word/Concepts/Working-with-Word/working-with-range-objects.md).


## Methods

- [BoldRun](Word.Selection.BoldRun.md)
- [Calculate](Word.Selection.Calculate.md)
- [ClearCharacterAllFormatting](Word.Selection.ClearCharacterAllFormatting.md)
- [ClearCharacterDirectFormatting](Word.Selection.ClearCharacterDirectFormatting.md)
- [ClearCharacterStyle](Word.Selection.ClearCharacterStyle.md)
- [ClearFormatting](Word.Selection.ClearFormatting.md)
- [ClearParagraphAllFormatting](Word.Selection.ClearParagraphAllFormatting.md)
- [ClearParagraphDirectFormatting](Word.Selection.ClearParagraphDirectFormatting.md)
- [ClearParagraphStyle](Word.Selection.ClearParagraphStyle.md)
- [Collapse](Word.Selection.Collapse.md)
- [ConvertToTable](Word.Selection.ConvertToTable.md)
- [Copy](Word.Selection.Copy.md)
- [CopyAsPicture](Word.Selection.CopyAsPicture.md)
- [CopyFormat](Word.Selection.CopyFormat.md)
- [CreateAutoTextEntry](Word.Selection.CreateAutoTextEntry.md)
- [CreateTextbox](Word.Selection.CreateTextbox.md)
- [Cut](Word.Selection.Cut.md)
- [Delete](Word.Selection.Delete.md)
- [DetectLanguage](Word.Selection.DetectLanguage.md)
- [EndKey](Word.Selection.EndKey.md)
- [EndOf](Word.Selection.EndOf.md)
- [EscapeKey](Word.Selection.EscapeKey.md)
- [Expand](Word.Selection.Expand.md)
- [ExportAsFixedFormat](Word.Selection.ExportAsFixedFormat.md)
- [ExportAsFixedFormat2](Word.Selection.ExportAsFixedFormat2.md)
- [Extend](Word.Selection.Extend.md)
- [GoTo](Word.Selection.GoTo.md)
- [GoToEditableRange](Word.Selection.GoToEditableRange.md)
- [GoToNext](Word.Selection.GoToNext.md)
- [GoToPrevious](Word.Selection.GoToPrevious.md)
- [HomeKey](Word.Selection.HomeKey.md)
- [InRange](Word.Selection.InRange.md)
- [InsertAfter](Word.Selection.InsertAfter.md)
- [InsertBefore](Word.Selection.InsertBefore.md)
- [InsertBreak](Word.Selection.InsertBreak.md)
- [InsertCaption](Word.Selection.InsertCaption.md)
- [InsertCells](Word.Selection.InsertCells.md)
- [InsertColumns](Word.Selection.InsertColumns.md)
- [InsertColumnsRight](Word.Selection.InsertColumnsRight.md)
- [InsertCrossReference](Word.Selection.InsertCrossReference.md)
- [InsertDateTime](Word.Selection.InsertDateTime.md)
- [InsertFile](Word.Selection.InsertFile.md)
- [InsertFormula](Word.Selection.InsertFormula.md)
- [InsertNewPage](Word.Selection.InsertNewPage.md)
- [InsertParagraph](Word.Selection.InsertParagraph.md)
- [InsertParagraphAfter](Word.Selection.InsertParagraphAfter.md)
- [InsertParagraphBefore](Word.Selection.InsertParagraphBefore.md)
- [InsertRows](Word.Selection.InsertRows.md)
- [InsertRowsAbove](Word.Selection.InsertRowsAbove.md)
- [InsertRowsBelow](Word.Selection.InsertRowsBelow.md)
- [InsertStyleSeparator](Word.Selection.InsertStyleSeparator.md)
- [InsertSymbol](Word.Selection.InsertSymbol.md)
- [InsertXML](Word.Selection.InsertXML.md)
- [InStory](Word.Selection.InStory.md)
- [IsEqual](Word.Selection.IsEqual.md)
- [ItalicRun](Word.Selection.ItalicRun.md)
- [LtrPara](Word.Selection.LtrPara.md)
- [LtrRun](Word.Selection.LtrRun.md)
- [Move](Word.Selection.Move.md)
- [MoveDown](Word.Selection.MoveDown.md)
- [MoveEnd](Word.Selection.MoveEnd.md)
- [MoveEndUntil](Word.Selection.MoveEndUntil.md)
- [MoveEndWhile](Word.Selection.MoveEndWhile.md)
- [MoveLeft](Word.Selection.MoveLeft.md)
- [MoveRight](Word.Selection.MoveRight.md)
- [MoveStart](Word.Selection.MoveStart.md)
- [MoveStartUntil](Word.Selection.MoveStartUntil.md)
- [MoveStartWhile](Word.Selection.MoveStartWhile.md)
- [MoveUntil](Word.Selection.MoveUntil.md)
- [MoveUp](Word.Selection.MoveUp.md)
- [MoveWhile](Word.Selection.MoveWhile.md)
- [Next](Word.Selection.Next.md)
- [NextField](Word.Selection.NextField.md)
- [NextRevision](Word.Selection.NextRevision.md)
- [NextSubdocument](Word.Selection.NextSubdocument.md)
- [Paste](Word.Selection.Paste.md)
- [PasteAndFormat](Word.Selection.PasteAndFormat.md)
- [PasteAppendTable](Word.Selection.PasteAppendTable.md)
- [PasteAsNestedTable](Word.Selection.PasteAsNestedTable.md)
- [PasteExcelTable](Word.Selection.PasteExcelTable.md)
- [PasteFormat](Word.Selection.PasteFormat.md)
- [PasteSpecial](Word.Selection.PasteSpecial.md)
- [Previous](Word.Selection.Previous.md)
- [PreviousField](Word.Selection.PreviousField.md)
- [PreviousRevision](Word.Selection.PreviousRevision.md)
- [PreviousSubdocument](Word.Selection.PreviousSubdocument.md)
- [ReadingModeGrowFont](Word.Selection.ReadingModeGrowFont.md)
- [ReadingModeShrinkFont](Word.Selection.ReadingModeShrinkFont.md)
- [RtlPara](Word.Selection.RtlPara.md)
- [RtlRun](Word.Selection.RtlRun.md)
- [Select](Word.Selection.Select.md)
- [SelectCell](Word.Selection.SelectCell.md)
- [SelectColumn](Word.Selection.SelectColumn.md)
- [SelectCurrentAlignment](Word.Selection.SelectCurrentAlignment.md)
- [SelectCurrentColor](Word.Selection.SelectCurrentColor.md)
- [SelectCurrentFont](Word.Selection.SelectCurrentFont.md)
- [SelectCurrentIndent](Word.Selection.SelectCurrentIndent.md)
- [SelectCurrentSpacing](Word.Selection.SelectCurrentSpacing.md)
- [SelectCurrentTabs](Word.Selection.SelectCurrentTabs.md)
- [SelectRow](Word.Selection.SelectRow.md)
- [SetRange](Word.Selection.SetRange.md)
- [Shrink](Word.Selection.Shrink.md)
- [ShrinkDiscontiguousSelection](Word.Selection.ShrinkDiscontiguousSelection.md)
- [Sort](Word.Selection.Sort.md)
- [SortAscending](Word.Selection.SortAscending.md)
- [SortByHeadings](Word.selection.sortbyheadings.md)
- [SortDescending](Word.Selection.SortDescending.md)
- [SplitTable](Word.Selection.SplitTable.md)
- [StartOf](Word.Selection.StartOf.md)
- [ToggleCharacterCode](Word.Selection.ToggleCharacterCode.md)
- [TypeBackspace](Word.Selection.TypeBackspace.md)
- [TypeParagraph](Word.Selection.TypeParagraph.md)
- [TypeText](Word.Selection.TypeText.md)
- [WholeStory](Word.Selection.WholeStory.md)

## Properties

- [Active](Word.Selection.Active.md)
- [Application](Word.Selection.Application.md)
- [BookmarkID](Word.Selection.BookmarkID.md)
- [Bookmarks](Word.Selection.Bookmarks.md)
- [Borders](Word.Selection.Borders.md)
- [Cells](Word.Selection.Cells.md)
- [Characters](Word.Selection.Characters.md)
- [ChildShapeRange](Word.Selection.ChildShapeRange.md)
- [Columns](Word.Selection.Columns.md)
- [ColumnSelectMode](Word.Selection.ColumnSelectMode.md)
- [Comments](Word.Selection.Comments.md)
- [Creator](Word.Selection.Creator.md)
- [Document](Word.Selection.Document.md)
- [Editors](Word.Selection.Editors.md)
- [End](Word.Selection.End.md)
- [EndnoteOptions](Word.Selection.EndnoteOptions.md)
- [Endnotes](Word.Selection.Endnotes.md)
- [EnhMetaFileBits](Word.Selection.EnhMetaFileBits.md)
- [ExtendMode](Word.Selection.ExtendMode.md)
- [Fields](Word.Selection.Fields.md)
- [Find](Word.Selection.Find.md)
- [FitTextWidth](Word.Selection.FitTextWidth.md)
- [Flags](Word.Selection.Flags.md)
- [Font](Word.Selection.Font.md)
- [FootnoteOptions](Word.Selection.FootnoteOptions.md)
- [Footnotes](Word.Selection.Footnotes.md)
- [FormattedText](Word.Selection.FormattedText.md)
- [FormFields](Word.Selection.FormFields.md)
- [Frames](Word.Selection.Frames.md)
- [HasChildShapeRange](Word.Selection.HasChildShapeRange.md)
- [HeaderFooter](Word.Selection.HeaderFooter.md)
- [HTMLDivisions](Word.Selection.HTMLDivisions.md)
- [Hyperlinks](Word.Selection.Hyperlinks.md)
- [Information](Word.Selection.Information.md)
- [InlineShapes](Word.Selection.InlineShapes.md)
- [IPAtEndOfLine](Word.Selection.IPAtEndOfLine.md)
- [IsEndOfRowMark](Word.Selection.IsEndOfRowMark.md)
- [LanguageDetected](Word.Selection.LanguageDetected.md)
- [LanguageID](Word.Selection.LanguageID.md)
- [LanguageIDFarEast](Word.Selection.LanguageIDFarEast.md)
- [LanguageIDOther](Word.Selection.LanguageIDOther.md)
- [NoProofing](Word.Selection.NoProofing.md)
- [OMaths](Word.Selection.OMaths.md)
- [Orientation](Word.Selection.Orientation.md)
- [PageSetup](Word.Selection.PageSetup.md)
- [ParagraphFormat](Word.Selection.ParagraphFormat.md)
- [Paragraphs](Word.Selection.Paragraphs.md)
- [Parent](Word.Selection.Parent.md)
- [PreviousBookmarkID](Word.Selection.PreviousBookmarkID.md)
- [Range](Word.Selection.Range.md)
- [Rows](Word.Selection.Rows.md)
- [Sections](Word.Selection.Sections.md)
- [Sentences](Word.Selection.Sentences.md)
- [Shading](Word.Selection.Shading.md)
- [ShapeRange](Word.Selection.ShapeRange.md)
- [Start](Word.Selection.Start.md)
- [StartIsActive](Word.Selection.StartIsActive.md)
- [StoryLength](Word.Selection.StoryLength.md)
- [StoryType](Word.Selection.StoryType.md)
- [Style](Word.Selection.Style.md)
- [Tables](Word.Selection.Tables.md)
- [Text](Word.Selection.Text.md)
- [TopLevelTables](Word.Selection.TopLevelTables.md)
- [Type](Word.Selection.Type.md)
- [WordOpenXML](Word.Selection.WordOpenXML.md)
- [Words](Word.Selection.Words.md)
- [XML](Word.Selection.XML.md)

## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
