---
title: View object (Word)
keywords: vbawd10.chm2469
f1_keywords:
- vbawd10.chm2469
ms.prod: word
api_name:
- Word.View
ms.assetid: 8bf5b26b-14c0-1985-65b2-3e034360baeb
ms.date: 08/15/2017
localization_priority: Normal
---


# View object (Word)

Contains the view attributes (such as show all, field shading, and table gridlines) for a window or pane.


## Remarks

Use the  **View** property to return the **View** object. The following example sets view options for the active window.


```vb
With ActiveDocument.ActiveWindow.View 
 .ShowAll = True 
 .TableGridlines = True 
 .WrapToWindow = False 
End With
```

Use the  **Type** property to change the view. The following example switches the active window to normal view.




```vb
ActiveDocument.ActiveWindow.View.Type = wdNormalView
```

Use the  **Percentage** property to change the size of the text on-screen. The following example enlarges the on-screen text to 120 percent.




```vb
ActiveDocument.ActiveWindow.View.Zoom.Percentage = 120
```

Use the  **SeekView** property to view comments, endnotes, footnotes, or the document header or footer. The following example displays the current footer in the active window in print layout view.




```vb
With ActiveDocument.ActiveWindow.View 
 .Type = wdPrintView 
 .SeekView = wdSeekCurrentPageFooter 
End With
```


## Methods



|Name|
|:-----|
|[CollapseAllHeadings](Word.view.collapseallheadings.md)|
|[CollapseOutline](Word.View.CollapseOutline.md)|
|[ExpandAllHeadings](Word.view.expandallheadings.md)|
|[ExpandOutline](Word.View.ExpandOutline.md)|
|[ForceLowresUpdate](overview/Word.md)|
|[ForceOffscreenUpdate](overview/Word.md)|
|[NextHeaderFooter](Word.View.NextHeaderFooter.md)|
|[PreviousHeaderFooter](Word.View.PreviousHeaderFooter.md)|
|[ShowAllHeadings](Word.View.ShowAllHeadings.md)|
|[ShowHeading](Word.View.ShowHeading.md)|

## Properties



|Name|
|:-----|
|[Application](Word.View.Application.md)|
|[ColumnWidth](Word.view.columnwidth.md)|
|[ConflictMode](Word.View.ConflictMode.md)|
|[Creator](Word.View.Creator.md)|
|[DisplayBackgrounds](Word.View.DisplayBackgrounds.md)|
|[DisplayPageBoundaries](Word.View.DisplayPageBoundaries.md)|
|[Draft](Word.View.Draft.md)|
|[FieldShading](Word.View.FieldShading.md)|
|[FullScreen](Word.View.FullScreen.md)|
|[Magnifier](Word.View.Magnifier.md)|
|[MailMergeDataView](Word.View.MailMergeDataView.md)|
|[MarkupMode](Word.View.MarkupMode.md)|
|[PageColor](Word.view.pagecolor.md)|
|[PageMovementType](Word.View.PageMovementType.md)|
|[Panning](Word.View.Panning.md)|
|[Parent](Word.View.Parent.md)|
|[ReadingLayout](Word.View.ReadingLayout.md)|
|[ReadingLayoutActualView](Word.View.ReadingLayoutActualView.md)|
|[ReadingLayoutTruncateMargins](Word.View.ReadingLayoutTruncateMargins.md)|
|[RevisionsBalloonShowConnectingLines](Word.View.RevisionsBalloonShowConnectingLines.md)|
|[RevisionsBalloonSide](Word.View.RevisionsBalloonSide.md)|
|[RevisionsBalloonWidth](Word.View.RevisionsBalloonWidth.md)|
|[RevisionsBalloonWidthType](Word.View.RevisionsBalloonWidthType.md)|
|[RevisionsFilter](Word.view.revisionsfilter.md)|
|[SeekView](Word.View.SeekView.md)|
|[ShadeEditableRanges](Word.View.ShadeEditableRanges.md)|
|[ShowAll](Word.View.ShowAll.md)|
|[ShowBookmarks](Word.View.ShowBookmarks.md)|
|[ShowComments](Word.View.ShowComments.md)|
|[ShowCropMarks](Word.View.ShowCropMarks.md)|
|[ShowDrawings](Word.View.ShowDrawings.md)|
|[ShowFieldCodes](Word.View.ShowFieldCodes.md)|
|[ShowFirstLineOnly](Word.View.ShowFirstLineOnly.md)|
|[ShowFormat](Word.View.ShowFormat.md)|
|[ShowFormatChanges](Word.View.ShowFormatChanges.md)|
|[ShowHiddenText](Word.View.ShowHiddenText.md)|
|[ShowHighlight](Word.View.ShowHighlight.md)|
|[ShowHyphens](Word.View.ShowHyphens.md)|
|[ShowInkAnnotations](Word.View.ShowInkAnnotations.md)|
|[ShowInsertionsAndDeletions](Word.View.ShowInsertionsAndDeletions.md)|
|[ShowMainTextLayer](Word.View.ShowMainTextLayer.md)|
|[ShowMarkupAreaHighlight](Word.View.ShowMarkupAreaHighlight.md)|
|[ShowObjectAnchors](Word.View.ShowObjectAnchors.md)|
|[ShowOptionalBreaks](Word.View.ShowOptionalBreaks.md)|
|[ShowOtherAuthors](Word.View.ShowOtherAuthors.md)|
|[ShowParagraphs](Word.View.ShowParagraphs.md)|
|[ShowPicturePlaceHolders](Word.View.ShowPicturePlaceHolders.md)|
|[ShowRevisionsAndComments](Word.View.ShowRevisionsAndComments.md)|
|[ShowSpaces](Word.View.ShowSpaces.md)|
|[ShowTabs](Word.View.ShowTabs.md)|
|[ShowTextBoundaries](Word.View.ShowTextBoundaries.md)|
|[ShowXMLMarkup](Word.View.ShowXMLMarkup.md)|
|[SplitSpecial](Word.View.SplitSpecial.md)|
|[TableGridlines](Word.View.TableGridlines.md)|
|[Type](Word.View.Type.md)|
|[WrapToWindow](Word.View.WrapToWindow.md)|
|[Zoom](Word.View.Zoom.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]