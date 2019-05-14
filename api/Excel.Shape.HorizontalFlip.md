---
title: Shape.HorizontalFlip property (Excel)
keywords: vbaxl10.chm636099
f1_keywords:
- vbaxl10.chm636099
ms.prod: excel
api_name:
- Excel.Shape.HorizontalFlip
ms.assetid: e9b64a81-3aef-5d42-0a65-5d03d564a71f
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.HorizontalFlip property (Excel)

**True** if the specified shape is flipped around the horizontal axis. Read-only **[MsoTriState](Office.MsoTriState.md)**.


## Syntax

_expression_.**HorizontalFlip**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


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