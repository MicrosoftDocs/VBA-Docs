---
title: ShapeRange.Flip method (Word)
keywords: vbawd10.chm162856975
f1_keywords:
- vbawd10.chm162856975
ms.prod: word
api_name:
- Word.ShapeRange.Flip
ms.assetid: 363c222b-f0fc-8d42-5b06-82ec607a00c7
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.Flip method (Word)

Flips a shape horizontally or vertically.


## Syntax

_expression_.**Flip** (_FlipCmd_)

_expression_ Required. A variable that represents a **[ShapeRange](Word.shaperange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FlipCmd_|Required| **MsoFlipCmd**|The flip orientation.|

## Example

This example adds a triangle to the active document, duplicates the triangle, and then flips the duplicate triangle vertically and makes it red.


```vb
Sub FlipShape() 
 With ActiveDocument.Shapes.AddShape( _ 
 Type:=msoShapeRightTriangle, Left:=150, _ 
 Top:=150, Width:=50, Height:=50).Duplicate 
 .Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 .Flip msoFlipVertical 
 End With 
End Sub
```


## See also


[ShapeRange Collection Object](Word.shaperange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]