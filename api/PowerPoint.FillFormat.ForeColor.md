---
title: FillFormat.ForeColor property (PowerPoint)
keywords: vbapp10.chm552012
f1_keywords:
- vbapp10.chm552012
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.ForeColor
ms.assetid: 3dc07a0f-d0bc-52c8-e06a-dd0315151742
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.ForeColor property (PowerPoint)

Returns or sets a **[ColorFormat](PowerPoint.ColorFormat.md)** object that represents the foreground color for the fill, line, or shadow. Read/write.


## Syntax

_expression_.**ForeColor**

_expression_ A variable that represents a **[FillFormat](powerpoint.fillformat.md)** object.


## Return value

ColorFormat


## Example

This example adds a rectangle to _myDocument_ and then sets the foreground color, background color, and gradient for the rectangle's fill.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes _
        .AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill

    .ForeColor.RGB = RGB(128, 0, 0)
    .BackColor.RGB = RGB(170, 170, 170)
    .TwoColorGradient msoGradientHorizontal, 1

End With
```

This example adds a patterned line to _myDocument_.




```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddLine(10, 100, 250, 0).Line

    .Weight = 6

    .ForeColor.RGB = RGB(0, 0, 255)

    .BackColor.RGB = RGB(128, 0, 0)

    .Pattern = msoPatternDarkDownwardDiagonal

End With
```


## See also


[FillFormat Object](PowerPoint.FillFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]