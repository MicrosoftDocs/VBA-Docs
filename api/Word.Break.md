---
title: Break object (Word)
keywords: vbawd10.chm3057
f1_keywords:
- vbawd10.chm3057
ms.prod: word
api_name:
- Word.Break
ms.assetid: 771ba998-c22e-3cf0-fab7-af9329793855
ms.date: 06/08/2017
localization_priority: Normal
---


# Break object (Word)

Represents individual page, column, and section breaks in a page. Use the **Break** object and the related methods and properties for programmatically defining page layout in a document.


## Remarks

Use the **[Item](Word.Breaks.Item.md)** method of the **[Breaks](Word.Breaks.md)** collection to return a specific **Break** object. The following example returns the first break in the first page of the active document.


```vb
Dim objBreak As Break 
 
Set objBreak = ActiveDocument.ActiveWindow _ 
 .Panes(1).Pages(1).Breaks.Item(1)
```

## Properties

- [Application](Word.Break.Application.md)
- [Creator](Word.Break.Creator.md)
- [PageIndex](Word.Break.PageIndex.md)
- [Parent](Word.Break.Parent.md)
- [Range](Word.Break.Range.md)

## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]