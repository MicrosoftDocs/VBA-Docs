---
title: ShapeRange.VerticalFlip property (Excel)
keywords: vbaxl10.chm640119
f1_keywords:
- vbaxl10.chm640119
ms.prod: excel
api_name:
- Excel.ShapeRange.VerticalFlip
ms.assetid: 43ecbc06-a16b-821f-b7c9-c66fcfad7a79
ms.date: 06/08/2017
---


# ShapeRange.VerticalFlip property (Excel)

 **True** if the specified shape is flipped around the vertical axis. Read-only **[MsoTriState](./Office.MsoTriState.md)**.


## Syntax

 _expression_. `VerticalFlip`

 _expression_ A variable that represents a [ShapeRange](./Excel.ShapeRange.md) object.


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


[ShapeRange Object](Excel.ShapeRange.md)

