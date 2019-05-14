---
title: ShapeRange.VerticalFlip property (Word)
keywords: vbawd10.chm162857085
f1_keywords:
- vbawd10.chm162857085
ms.prod: word
api_name:
- Word.ShapeRange.VerticalFlip
ms.assetid: f4dc248c-3ffa-e7e3-8ca9-9f6afc8be832
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.VerticalFlip property (Word)

 **True** if the specified shape is flipped around the vertical axis. Read-only **MsoTriState**.


## Syntax

_expression_.**VerticalFlip**

_expression_ Required. A variable that represents a **[ShapeRange](Word.shaperange.md)** object.


## Example

This example restores each shape on _myDocument_ to its original state if it has been flipped horizontally or vertically.


```vb
For Each s In ActiveDocument.Range.ShapeRange 
 If s.HorizontalFlip Then s.Flip msoFlipHorizontal 
 If s.VerticalFlip Then s.Flip msoFlipVertical 
Next
```


## See also


[ShapeRange Collection Object](Word.shaperange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]