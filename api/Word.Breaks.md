---
title: Breaks object (Word)
keywords: vbawd10.chm777
f1_keywords:
- vbawd10.chm777
ms.prod: word
api_name:
- Word.Breaks
ms.assetid: c6a9213d-c4bc-a9a6-3889-a32446f1fc25
ms.date: 06/08/2017
localization_priority: Normal
---


# Breaks object (Word)

A collection of page, column, or section breaks in a page. Use the  **Breaks** collection and the related objects and properties to programmatically define page layout in a document.


## Remarks

Use the  **[Breaks](Word.Page.Breaks.md)** property to return a **Breaks** collection. The following example returns the breaks in the first page of the active document.


```vb
Dim objBreaks As Breaks 
 
Set objBreaks = ActiveDocument.ActiveWindow _ 
 .Panes(1).Pages(1).Breaks
```

## Methods

- [Item](Word.Breaks.Item.md)

## Properties

- [Application](Word.Breaks.Application.md)
- [Count](Word.Breaks.Count.md)
- [Creator](Word.Breaks.Creator.md)
- [Parent](Word.Breaks.Parent.md)

## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]