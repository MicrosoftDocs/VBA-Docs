---
title: ShadowFormat object (Excel)
keywords: vbaxl10.chm114000
f1_keywords:
- vbaxl10.chm114000
ms.prod: excel
api_name:
- Excel.ShadowFormat
ms.assetid: 2566c68e-f8d6-badc-3ce9-b6ae5f9c1cc2
ms.date: 04/02/2019
localization_priority: Normal
---


# ShadowFormat object (Excel)

Represents shadow formatting for a shape.


## Remarks

Use the **[Shadow](Excel.Shape.Shadow.md)** property of the **Shape** object to return a **ShadowFormat** object.


## Example

The following example adds a shadowed rectangle to _myDocument_. The semi-transparent, blue shadow is offset 5 points to the right of the rectangle and 3 points above it.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 50, 50, 100, 200).Shadow 
 .ForeColor.RGB = RGB(0, 0, 128) 
 .OffsetX = 5 
 .OffsetY = -3 
 .Transparency = 0.5 
 .Visible = True 
End With
```

## Methods

- [IncrementOffsetX](Excel.ShadowFormat.IncrementOffsetX.md)
- [IncrementOffsetY](Excel.ShadowFormat.IncrementOffsetY.md)

## Properties

- [Application](Excel.ShadowFormat.Application.md)
- [Blur](Excel.ShadowFormat.Blur.md)
- [Creator](Excel.ShadowFormat.Creator.md)
- [ForeColor](Excel.ShadowFormat.ForeColor.md)
- [Obscured](Excel.ShadowFormat.Obscured.md)
- [OffsetX](Excel.ShadowFormat.OffsetX.md)
- [OffsetY](Excel.ShadowFormat.OffsetY.md)
- [Parent](Excel.ShadowFormat.Parent.md)
- [RotateWithShape](Excel.ShadowFormat.RotateWithShape.md)
- [Size](Excel.ShadowFormat.Size.md)
- [Style](Excel.ShadowFormat.Style.md)
- [Transparency](Excel.ShadowFormat.Transparency.md)
- [Type](Excel.ShadowFormat.Type.md)
- [Visible](Excel.ShadowFormat.Visible.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]