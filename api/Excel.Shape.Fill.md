---
title: Shape.Fill property (Excel)
keywords: vbaxl10.chm636096
f1_keywords:
- vbaxl10.chm636096
ms.prod: excel
api_name:
- Excel.Shape.Fill
ms.assetid: b533b463-51c5-f59e-c3ba-cfe7512daa53
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.Fill property (Excel)

Returns a **[FillFormat](Excel.FillFormat.md)** object for a specified shape or a **[ChartFillFormat](Excel.ChartFillFormat.md)** object for a specified chart that contains fill formatting properties for the shape or chart. Read-only.


## Syntax

_expression_.**Fill**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


## Example

This example adds a rectangle to _myDocument_ and then sets the foreground color, background color, and gradient for the rectangle's fill.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
        90, 90, 90, 50).Fill 
    .ForeColor.RGB = RGB(128, 0, 0) 
    .BackColor.RGB = RGB(170, 170, 170) 
    .TwoColorGradient msoGradientHorizontal, 1 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
