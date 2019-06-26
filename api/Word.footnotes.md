---
title: Footnotes object (Word)
ms.prod: word
ms.assetid: d46a0972-2784-4814-d547-30122a35cdc1
ms.date: 06/08/2017
localization_priority: Normal
---


# Footnotes object (Word)

A collection of  **Footnote** objects that represent all the footnotes in a selection, range, or document.


## Remarks

Use the  **Footnotes** property to return the **Footnotes** collection. The following example changes all of the footnotes in the active document to endnotes.


```vb
ActiveDocument.Footnotes.SwapWithEndnotes
```

Use the  **Add** method to add a footnote to the **Footnotes** collection. The following example adds a footnote immediately after the selection.




```vb
Selection.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Footnotes.Add Range:=Selection.Range , _ 
 Text:="The Willow Tree, (Lone Creek Press, 1996)."
```

Use  **Footnotes** (_index_), where _index_ is the index number, to return a single **[Footnote](Word.Footnote.md)** object. The index number represents the position of the footnote in the selection, range, or document. The following example applies red formatting to the first footnote in the selection.




```vb
If Selection.Footnotes.Count >= 1 Then 
 Selection.Footnotes(1).Reference.Font.ColorIndex = wdRed 
End If
```


> [!NOTE] 
> Footnotes positioned at the end of a document or section are considered endnotes and are included in the  **[Endnotes](Word.endnotes.md)** collection.


## Methods



|Name|
|:-----|
|[Add](Word.Footnotes.Add.md)|
|[Convert](Word.Footnotes.Convert.md)|
|[Item](Word.Footnotes.Item.md)|
|[ResetContinuationNotice](Word.Footnotes.ResetContinuationNotice.md)|
|[ResetContinuationSeparator](Word.Footnotes.ResetContinuationSeparator.md)|
|[ResetSeparator](Word.Footnotes.ResetSeparator.md)|
|[SwapWithEndnotes](Word.Footnotes.SwapWithEndnotes.md)|

## Properties



|Name|
|:-----|
|[Application](Word.Footnotes.Application.md)|
|[ContinuationNotice](Word.Footnotes.ContinuationNotice.md)|
|[ContinuationSeparator](Word.Footnotes.ContinuationSeparator.md)|
|[Count](Word.Footnotes.Count.md)|
|[Creator](Word.Footnotes.Creator.md)|
|[Location](Word.Footnotes.Location.md)|
|[NumberingRule](Word.Footnotes.NumberingRule.md)|
|[NumberStyle](Word.Footnotes.NumberStyle.md)|
|[Parent](Word.Footnotes.Parent.md)|
|[Separator](Word.Footnotes.Separator.md)|
|[StartingNumber](Word.Footnotes.StartingNumber.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]