---
title: LineFormat object (Excel)
keywords: vbaxl10.chm110000
f1_keywords:
- vbaxl10.chm110000
ms.prod: excel
api_name:
- Excel.LineFormat
ms.assetid: 13eca34b-adf7-ddd3-8c73-cc8b508c624a
ms.date: 03/30/2019
localization_priority: Normal
---


# LineFormat object (Excel)

Represents line and arrowhead formatting.


## Remarks

For a line, the **LineFormat** object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.


## Example

Use the **[Line](Excel.Shape.Line.md)** property of the **Shape** object to return a **LineFormat** object. 

The following example adds a blue, dashed line to _myDocument_. There's a short, narrow oval at the line's starting point and a long, wide triangle at its end point.

```vb
Set myDocument = Worksheets(1) 
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

- [Application](Excel.LineFormat.Application.md)
- [BackColor](Excel.LineFormat.BackColor.md)
- [BeginArrowheadLength](Excel.LineFormat.BeginArrowheadLength.md)
- [BeginArrowheadStyle](Excel.LineFormat.BeginArrowheadStyle.md)
- [BeginArrowheadWidth](Excel.LineFormat.BeginArrowheadWidth.md)
- [Creator](Excel.LineFormat.Creator.md)
- [DashStyle](Excel.LineFormat.DashStyle.md)
- [EndArrowheadLength](Excel.LineFormat.EndArrowheadLength.md)
- [EndArrowheadStyle](Excel.LineFormat.EndArrowheadStyle.md)
- [EndArrowheadWidth](Excel.LineFormat.EndArrowheadWidth.md)
- [ForeColor](Excel.LineFormat.ForeColor.md)
- [InsetPen](Excel.LineFormat.InsetPen.md)
- [Parent](Excel.LineFormat.Parent.md)
- [Pattern](Excel.LineFormat.Pattern.md)
- [Style](Excel.LineFormat.Style.md)
- [Transparency](Excel.LineFormat.Transparency.md)
- [Visible](Excel.LineFormat.Visible.md)
- [Weight](Excel.LineFormat.Weight.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
