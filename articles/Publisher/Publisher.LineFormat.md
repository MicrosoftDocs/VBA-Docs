---
title: LineFormat Object (Publisher)
keywords: vbapb10.chm3473407
f1_keywords:
- vbapb10.chm3473407
ms.prod: publisher
api_name:
- Publisher.LineFormat
ms.assetid: 9c973f5a-b2d2-78b1-24c3-350f1ba4c2ab
ms.date: 06/08/2017
---


# LineFormat Object (Publisher)

Represents line and arrowhead formatting. For a line, the  **LineFormat** object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.
 


## Example

Use the  **[Line](Publisher.Shape.Line.md)** property to return a **LineFormat** object. The following example adds a blue, dashed line to the active document. There is a short, narrow oval at the line's starting point and a long, wide triangle at its endpoint.
 

 

```
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



|**Name**|
|:-----|
|[PresetGradient](lineformat-presetgradient-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](Publisher.LineFormat.Application.md)|
|[BackColor](Publisher.LineFormat.BackColor.md)|
|[BeginArrowheadLength](Publisher.LineFormat.BeginArrowheadLength.md)|
|[BeginArrowheadStyle](Publisher.LineFormat.BeginArrowheadStyle.md)|
|[BeginArrowheadWidth](Publisher.LineFormat.BeginArrowheadWidth.md)|
|[CapStyle](lineformat-capstyle-property-publisher.md)|
|[DashStyle](Publisher.LineFormat.DashStyle.md)|
|[EndArrowheadLength](Publisher.LineFormat.EndArrowheadLength.md)|
|[EndArrowheadStyle](Publisher.LineFormat.EndArrowheadStyle.md)|
|[EndArrowheadWidth](Publisher.LineFormat.EndArrowheadWidth.md)|
|[ForeColor](Publisher.LineFormat.ForeColor.md)|
|[GradientAngle](lineformat-gradientangle-property-publisher.md)|
|[GradientColorType](lineformat-gradientcolortype-property-publisher.md)|
|[GradientStyle](lineformat-gradientstyle-property-publisher.md)|
|[GradientVariant](lineformat-gradientvariant-property-publisher.md)|
|[InsetPen](Publisher.LineFormat.InsetPen.md)|
|[JoinStyle](lineformat-joinstyle-property-publisher.md)|
|[Parent](Publisher.LineFormat.Parent.md)|
|[Pattern](Publisher.LineFormat.Pattern.md)|
|[PresetGradientType](lineformat-presetgradienttype-property-publisher.md)|
|[Style](Publisher.LineFormat.Style.md)|
|[Transparency](lineformat-transparency-property-publisher.md)|
|[Type](lineformat-type-property-publisher.md)|
|[Visible](Publisher.LineFormat.Visible.md)|
|[Weight](Publisher.LineFormat.Weight.md)|

