---
title: ShapeRange.VerticalFlip property (Excel)
keywords: vbaxl10.chm640119
f1_keywords:
- vbaxl10.chm640119
ms.prod: excel
api_name:
- Excel.ShapeRange.VerticalFlip
ms.assetid: 43ecbc06-a16b-821f-b7c9-c66fcfad7a79
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.VerticalFlip property (Excel)

**True** if the specified shape is flipped around the vertical axis. Read-only **[MsoTriState](Office.MsoTriState.md)**.


## Syntax

_expression_.**VerticalFlip**

_expression_ A variable that represents a **[ShapeRange](Excel.shaperange.md)** object.


## Example

This example restores each shape on _myDocument_ to its original state if it has been flipped horizontally or vertically.

```vb
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
    If s.HorizontalFlip Then s.Flip msoFlipHorizontal 
    If s.VerticalFlip Then s.Flip msoFlipVertical 
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]