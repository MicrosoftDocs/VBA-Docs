---
title: Shape.VerticalFlip property (Word)
keywords: vbawd10.chm161480829
f1_keywords:
- vbawd10.chm161480829
ms.prod: word
api_name:
- Word.Shape.VerticalFlip
ms.assetid: f14d27b2-99f5-ddf5-a6b9-4163c20c0715
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.VerticalFlip property (Word)

 **True** if the specified shape is flipped around the vertical axis. Read-only **MsoTriState**.


## Syntax

_expression_.**VerticalFlip**

_expression_ Required. A variable that represents a **[Shape](Word.Shape.md)** object.


## Example

This example restores each shape on _myDocument_ to its original state if it has been flipped horizontally or vertically.


```vb
For Each s In ActiveDocument.Shapes 
 If s.HorizontalFlip Then s.Flip msoFlipHorizontal 
 If s.VerticalFlip Then s.Flip msoFlipVertical 
Next
```


## See also


[Shape Object](Word.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]