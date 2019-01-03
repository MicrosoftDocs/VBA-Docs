---
title: Shape.VerticalFlip property (Excel)
keywords: vbaxl10.chm636112
f1_keywords:
- vbaxl10.chm636112
ms.prod: excel
api_name:
- Excel.Shape.VerticalFlip
ms.assetid: 3b50edac-a167-8e07-3286-6ced14bb715d
ms.date: 06/08/2017
---


# Shape.VerticalFlip property (Excel)

 **True** if the specified shape is flipped around the vertical axis. Read-only **[MsoTriState](Office.MsoTriState.md)**.


## Syntax

 _expression_. `VerticalFlip`

 _expression_ A variable that represents a [Shape](./Excel.Shape.md) object.


## Example

This example restores each shape on  `myDocument` to its original state if it's been flipped horizontally or vertically.


```vb
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
    If s.HorizontalFlip Then s.Flip msoFlipHorizontal 
    If s.VerticalFlip Then s.Flip msoFlipVertical 
Next
```


## See also


[Shape Object](Excel.Shape.md)

