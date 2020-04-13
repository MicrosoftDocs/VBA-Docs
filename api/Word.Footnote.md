---
title: Footnote object (Word)
keywords: vbawd10.chm2367
f1_keywords:
- vbawd10.chm2367
ms.prod: word
api_name:
- Word.Footnote
ms.assetid: 877340c4-14f9-4560-eaf8-2c6482a1ade8
ms.date: 06/08/2017
localization_priority: Normal
---


# Footnote object (Word)

Represents a footnote positioned at the bottom of the page or beneath text. The **Footnote** object is a member of the **Footnotes** collection. The **[Footnotes](Word.footnotes.md)** collection represents the footnotes in a selection, range, or document.


## Remarks

Use  **Footnotes** (Index), where Index is the index number, to return a single **Footnote** object. The index number represents the position of the footnote in the selection, range, or document. The following example applies red formatting to the first footnote in the selection.


```vb
If Selection.Footnotes.Count >= 1 Then 
 Selection.Footnotes(1).Reference.Font.ColorIndex = wdRed 
End If
```

Use the **Add** method to add a footnote to the **[Footnotes](Word.footnotes.md)** collection. The following example inserts an automatically numbered footnote immediately after the selection.




```vb
Selection.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Footnotes.Add Range:=Selection.Range , _ 
 Text:="The Willow Tree, (Lone Creek Press, 1996)."
```


> [!NOTE] 
> Footnotes positioned at the end of a document or section are considered endnotes and are included in the **[Endnotes](Word.endnotes.md)** collection.


## Methods



|Name|
|:-----|
|[Delete](Word.Footnote.Delete.md)|

## Properties



|Name|
|:-----|
|[Application](Word.Footnote.Application.md)|
|[Creator](Word.Footnote.Creator.md)|
|[Index](Word.Footnote.Index.md)|
|[Parent](Word.Footnote.Parent.md)|
|[Range](Word.Footnote.Range.md)|
|[Reference](Word.Footnote.Reference.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]