---
title: LineFormat object (PowerPoint)
keywords: vbapp10.chm553000
f1_keywords:
- vbapp10.chm553000
ms.prod: powerpoint
api_name:
- PowerPoint.LineFormat
ms.assetid: 11c955d5-bbda-d99f-cec9-fc6187450a12
ms.date: 06/08/2017
localization_priority: Normal
---


# LineFormat object (PowerPoint)

Represents line and arrowhead formatting. For a line, the  **LineFormat** object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.


## Example

Use the  **Line** property to return a **LineFormat** object. The following example adds a blue, dashed line to _myDocument_. There's a short, narrow oval at the line's starting point and a long, wide triangle at its endpoint.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddLine(100, 100, 200, 300).Line

    .DashStyle = msoLineDashDotDot

    .ForeColor.RGB = RGB(50, 0, 128)

    .BeginArrowheadLength = msoArrowheadShort

    .BeginArrowheadStyle = msoArrowheadOval

    .BeginArrowheadWidth = msoArrowheadNarrow

    .EndArrowheadLength = msoArrowheadLong

    .EndArrowheadStyle = msoArrowheadTriangle

    .EndArrowheadWidth = msoArrowheadWide

End With
```


## Properties



|Name|
|:-----|
|[Application](PowerPoint.LineFormat.Application.md)|
|[BackColor](PowerPoint.LineFormat.BackColor.md)|
|[BeginArrowheadLength](PowerPoint.LineFormat.BeginArrowheadLength.md)|
|[BeginArrowheadStyle](PowerPoint.LineFormat.BeginArrowheadStyle.md)|
|[BeginArrowheadWidth](PowerPoint.LineFormat.BeginArrowheadWidth.md)|
|[Creator](PowerPoint.LineFormat.Creator.md)|
|[DashStyle](PowerPoint.LineFormat.DashStyle.md)|
|[EndArrowheadLength](PowerPoint.LineFormat.EndArrowheadLength.md)|
|[EndArrowheadStyle](PowerPoint.LineFormat.EndArrowheadStyle.md)|
|[EndArrowheadWidth](PowerPoint.LineFormat.EndArrowheadWidth.md)|
|[ForeColor](PowerPoint.LineFormat.ForeColor.md)|
|[InsetPen](PowerPoint.LineFormat.InsetPen.md)|
|[Parent](PowerPoint.LineFormat.Parent.md)|
|[Pattern](PowerPoint.LineFormat.Pattern.md)|
|[Style](PowerPoint.LineFormat.Style.md)|
|[Transparency](PowerPoint.LineFormat.Transparency.md)|
|[Visible](PowerPoint.LineFormat.Visible.md)|
|[Weight](PowerPoint.LineFormat.Weight.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]