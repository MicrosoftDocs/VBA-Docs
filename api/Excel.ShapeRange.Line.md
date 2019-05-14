---
title: ShapeRange.Line property (Excel)
keywords: vbaxl10.chm640108
f1_keywords:
- vbaxl10.chm640108
ms.prod: excel
api_name:
- Excel.ShapeRange.Line
ms.assetid: 7504afaa-0ddd-6ae8-4653-fddc0af9ede7
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.Line property (Excel)

Returns a **[LineFormat](Excel.LineFormat.md)** object that contains line formatting properties for the specified shape. (For a line, the **LineFormat** object represents the line itself; for a shape with a border, the **LineFormat** object represents the border). Read-only.


## Syntax

_expression_.**Line**

_expression_ A variable that represents a **[ShapeRange](Excel.shaperange.md)** object.


## Example

This example adds a blue dashed line to _myDocument_.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddLine(10, 10, 250, 250).Line 
 .DashStyle = msoLineDashDotDot 
 .ForeColor.RGB = RGB(50, 0, 128) 
End With
```

<br/>

This example adds a cross to _myDocument_ and then sets its border to be 8 points thick and red.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeCross, 10, 10, 50, 70).Line 
 .Weight = 8 
 .ForeColor.RGB = RGB(255, 0, 0) 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]