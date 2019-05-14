---
title: InlineShape.ScaleWidth property (Word)
keywords: vbawd10.chm162005003
f1_keywords:
- vbawd10.chm162005003
ms.prod: word
api_name:
- Word.InlineShape.ScaleWidth
ms.assetid: 64a22966-2516-758a-1f83-d4eaf09e0040
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShape.ScaleWidth property (Word)

Scales the width of the specified inline shape relative to its original size. Read/write  **Single**.


## Syntax

_expression_.**ScaleWidth**

_expression_ Required. A variable that represents an '[InlineShape](Word.InlineShape.md)' object.


## Example

This example sets the height and width of the first inline shape in the active document to 150 percent of the shape's original height and width.


```vb
With ActiveDocument.InlineShapes(1) 
 .ScaleHeight = 150 
 .ScaleWidth = 150 
End With
```


## See also


[InlineShape Object](Word.InlineShape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]