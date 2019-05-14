---
title: ShapeRange.HorizontalFlip property (Excel)
keywords: vbaxl10.chm640106
f1_keywords:
- vbaxl10.chm640106
ms.prod: excel
api_name:
- Excel.ShapeRange.HorizontalFlip
ms.assetid: 3b5f3755-987c-cd48-44a2-8be8bdd886dd
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.HorizontalFlip property (Excel)

**True** if the specified shape is flipped around the horizontal axis. Read-only **[MsoTriState](Office.MsoTriState.md)**.


## Syntax

_expression_.**HorizontalFlip**

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