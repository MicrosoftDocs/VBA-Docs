---
title: LineFormat object (Publisher)
keywords: vbapb10.chm3473407
f1_keywords:
- vbapb10.chm3473407
ms.prod: publisher
api_name:
- Publisher.LineFormat
ms.assetid: 9c973f5a-b2d2-78b1-24c3-350f1ba4c2ab
ms.date: 05/31/2019
localization_priority: Normal
---


# LineFormat object (Publisher)

Represents line and arrowhead formatting. For a line, the **LineFormat** object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.

## Remarks

Use the **[Shape.Line](Publisher.Shape.Line.md)** property to return a **LineFormat** object. 
 
## Example

The following example adds a blue, dashed line to the active document. There is a short, narrow oval at the line's starting point and a long, wide triangle at its endpoint.
 
```vb
Sub FormatLine() 
 With ActiveDocument.Pages(1).Shapes.AddLine(BeginX:=100, _ 
 BeginY:=100, EndX:=200, EndY:=300).Line 
 .DashStyle = msoLineDashDotDot 
 .ForeColor.RGB = RGB(50, 0, 128) 
 .BeginArrowheadLength = msoArrowheadShort 
 .BeginArrowheadStyle = msoArrowheadOval 
 .BeginArrowheadWidth = msoArrowheadNarrow 
 .EndArrowheadLength = msoArrowheadLong 
 .EndArrowheadStyle = msoArrowheadTriangle 
 .EndArrowheadWidth = msoArrowheadWide 
 End With 
End Sub
```


## Methods

- [PresetGradient](Publisher.lineformat.presetgradient.md)

## Properties

- [Application](Publisher.LineFormat.Application.md)
- [BackColor](Publisher.LineFormat.BackColor.md)
- [BeginArrowheadLength](Publisher.LineFormat.BeginArrowheadLength.md)
- [BeginArrowheadStyle](Publisher.LineFormat.BeginArrowheadStyle.md)
- [BeginArrowheadWidth](Publisher.LineFormat.BeginArrowheadWidth.md)
- [CapStyle](Publisher.lineformat.capstyle.md)
- [DashStyle](Publisher.LineFormat.DashStyle.md)
- [EndArrowheadLength](Publisher.LineFormat.EndArrowheadLength.md)
- [EndArrowheadStyle](Publisher.LineFormat.EndArrowheadStyle.md)
- [EndArrowheadWidth](Publisher.LineFormat.EndArrowheadWidth.md)
- [ForeColor](Publisher.LineFormat.ForeColor.md)
- [GradientAngle](Publisher.lineformat.gradientangle.md)
- [GradientColorType](Publisher.lineformat.gradientcolortype.md)
- [GradientStyle](Publisher.lineformat.gradientstyle.md)
- [GradientVariant](Publisher.lineformat.gradientvariant.md)
- [InsetPen](Publisher.LineFormat.InsetPen.md)
- [JoinStyle](Publisher.lineformat.joinstyle.md)
- [Parent](Publisher.LineFormat.Parent.md)
- [Pattern](Publisher.LineFormat.Pattern.md)
- [PresetGradientType](Publisher.lineformat.presetgradienttype.md)
- [Style](Publisher.LineFormat.Style.md)
- [Transparency](Publisher.lineformat.transparency.md)
- [Type](Publisher.lineformat.type.md)
- [Visible](Publisher.LineFormat.Visible.md)
- [Weight](Publisher.LineFormat.Weight.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]