---
title: Shape.Flip method (Word)
keywords: vbawd10.chm161480717
f1_keywords:
- vbawd10.chm161480717
ms.prod: word
api_name:
- Word.Shape.Flip
ms.assetid: 6da14b0b-9e75-7891-dd92-db771cc242a4
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Flip method (Word)

Flips a shape horizontally or vertically.


## Syntax

_expression_.**Flip** (_FlipCmd_)

_expression_ Required. A variable that represents a **[Shape](Word.Shape.md)** object.


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


[Shape Object](Word.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]