---
title: ShapeRange.Flip method (Excel)
keywords: vbaxl10.chm640082
f1_keywords:
- vbaxl10.chm640082
ms.prod: excel
api_name:
- Excel.ShapeRange.Flip
ms.assetid: 65f8066d-a522-ac67-662b-8c31a47fb725
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.Flip method (Excel)

Flips the specified shape around its horizontal or vertical axis.


## Syntax

_expression_. `Flip`( `_FlipCmd_` )

_expression_ A variable that represents a [ShapeRange](./Excel.ShapeRange.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FlipCmd_|Required| **[MsoFlipCmd](Office.MsoFlipCmd.md)**|Specifies whether the shape is to be flipped horizontally or vertically.|

## Example

This example adds a triangle to  `myDocument`, duplicates the triangle, and then flips the duplicate triangle vertically and makes it red.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRightTriangle, _ 
        10, 10, 50, 50).Duplicate 
    .Fill.ForeColor.RGB = RGB(255, 0, 0) 
    .Flip msoFlipVertical 
End With
```


## See also


[ShapeRange Object](Excel.ShapeRange.md)

