---
title: Shape.Line property (PowerPoint)
keywords: vbapp10.chm547027
f1_keywords:
- vbapp10.chm547027
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Line
ms.assetid: edb5f40e-8b1e-fd3f-33da-0d4f1d465525
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Line property (PowerPoint)

Returns a **[LineFormat](PowerPoint.LineFormat.md)** object that contains line formatting properties for the specified shape. (For a line, the **LineFormat** object represents the line itself; for a shape with a border, the **LineFormat** object represents the border.) Read-only.


## Syntax

_expression_.**Line**

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Return value

LineFormat


## Example

This example adds a blue dashed line to _myDocument_.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddLine(10, 10, 250, 250).Line

    .DashStyle = msoLineDashDotDot

    .ForeColor.RGB = RGB(50, 0, 128)

End With
```

This example adds a cross to the first slide and then sets its border to be 8 points thick and red.




```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddShape(msoShapeCross, 10, 10, 50, 70).Line

    .Weight = 8

    .ForeColor.RGB = RGB(255, 0, 0)

End With
```


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]