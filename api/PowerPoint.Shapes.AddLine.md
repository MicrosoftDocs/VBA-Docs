---
title: Shapes.AddLine method (PowerPoint)
keywords: vbapp10.chm543009
f1_keywords:
- vbapp10.chm543009
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddLine
ms.assetid: 9dbe640b-5ba4-a620-d3c6-4a2d0cc2bc27
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddLine method (PowerPoint)

Creates a line. Returns a **[Shape](PowerPoint.Shape.md)** object that represents the new line.


## Syntax

_expression_. `AddLine`( `_BeginX_`, `_BeginY_`, `_EndX_`, `_EndY_` )

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _BeginX_|Required|**Single**|The horizontal position, measured in points, of the line's starting point relative to the left edge of the slide.|
| _BeginY_|Required|**Single**|The vertical position, measured in points, of the line's starting point relative to the top edge of the slide.|
| _EndX_|Required|**Single**|The horizontal position, measured in points, of the line's ending point relative to the left edge of the slide.|
| _EndY_|Required|**Single**|The vertical position, measured in points, of the line's ending point relative to the top edge of the slide.|

## Return value

Shape


## Example

This example adds a blue dashed line to myDocument.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddLine(BeginX:=10, BeginY:=10, _
        EndX:=250, EndY:=250).Line
    .DashStyle = msoLineDashDotDot
    .ForeColor.RGB = RGB(50, 0, 128)
End With
```


## See also


[Shapes Object](PowerPoint.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]