---
title: TextRange.BoundWidth Property (PowerPoint)
keywords: vbapp10.chm569008
f1_keywords:
- vbapp10.chm569008
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.BoundWidth
ms.assetid: 409d1c66-8956-cdd0-2328-f1cbe584f778
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRange.BoundWidth Property (PowerPoint)

Returns the width (in points) of the text bounding box for the specified text frame. Read-only.


## Syntax

 _expression_. `BoundWidth`

 _expression_ A variable that represents a [TextRange](./PowerPoint.TextRange.md) object.


## Return value

Single


## Example

This example adds a rounded rectangle to slide one in the active presentation. The rectangle has the same dimensions as the text bounding box for shape one.


```vb
With Application.ActivePresentation.Slides(1).Shapes
    Set tr = .Item(1).TextFrame.TextRange
    Set roundRect = .AddShape(msoShapeRoundedRectangle, _
        tr.BoundLeft, tr.BoundTop, tr.BoundWidth, tr.BoundHeight)
End With

With roundRect.Fill
    .ForeColor.RGB = RGB(255, 0, 128)
    .Transparency = 0.75
End With
```


## See also


[TextRange Object](PowerPoint.TextRange.md)

