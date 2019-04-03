---
title: Rectangles object (Word)
ms.prod: word
api_name:
- Word.Rectangles
ms.assetid: c1de5e7f-13b1-e35a-d9f1-9a8f1246e2e5
ms.date: 06/08/2017
localization_priority: Normal
---


# Rectangles object (Word)

A collection of  **Rectangle** objects in a page that represent portions of text and graphics. Use the **Rectangles** collection and related objects and properties for programmatically defining page layout in a document.


## Remarks

Use the  **Rectangles** property to return a **Rectangles** collection. The following example returns the **Rectangles** collection for the first page in the active document.


```vb
Dim objRectangles As Rectangles 
 
Set objRectangles = ActiveDocument.ActiveWindow _ 
 .Panes(1).Pages(1).Rectangles
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]