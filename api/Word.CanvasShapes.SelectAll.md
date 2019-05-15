---
title: CanvasShapes.SelectAll method (Word)
keywords: vbawd10.chm7536662
f1_keywords:
- vbawd10.chm7536662
ms.prod: word
api_name:
- Word.CanvasShapes.SelectAll
ms.assetid: c11c375a-8fb3-535d-b49a-2262560021dd
ms.date: 06/08/2017
localization_priority: Normal
---


# CanvasShapes.SelectAll method (Word)

Selects all the shapes in a canvas.


## Syntax

_expression_.**SelectAll**

_expression_ Required. A variable that represents a **[CanvasShapes](Word.CanvasShapes.md)** object.


## Remarks

This method doesn't select  **InlineShape** objects.


## Example

This example selects and deletes all the shapes inside the first canvas of the active document.


```vb
Sub SelectCanvasShapes() 
 Dim s As Shape 
 Set s = ActiveDocument.Shapes.Range(1) 
 s.CanvasItems.SelectAll 
 Selection.Delete 
End Sub
```


## See also


[CanvasShapes Collection](Word.CanvasShapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]