---
title: Shape.Flip method (Excel)
keywords: vbaxl10.chm636077
f1_keywords:
- vbaxl10.chm636077
ms.prod: excel
api_name:
- Excel.Shape.Flip
ms.assetid: 6ba41c89-878e-d9e1-5594-0cf45411b608
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.Flip method (Excel)

Flips the specified shape around its horizontal or vertical axis.


## Syntax

_expression_.**Flip** (_FlipCmd_)

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FlipCmd_|Required| **[MsoFlipCmd](Office.MsoFlipCmd.md)**|Specifies whether the shape is to be flipped horizontally or vertically.|

## Example

This example adds a triangle to _myDocument_, duplicates the triangle, and then flips the duplicate triangle vertically and makes it red.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRightTriangle, _ 
        10, 10, 50, 50).Duplicate 
    .Fill.ForeColor.RGB = RGB(255, 0, 0) 
    .Flip msoFlipVertical 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]