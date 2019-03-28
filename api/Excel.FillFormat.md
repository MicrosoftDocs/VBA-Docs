---
title: FillFormat object (Excel)
keywords: vbaxl10.chm115000
f1_keywords:
- vbaxl10.chm115000
ms.prod: excel
api_name:
- Excel.FillFormat
ms.assetid: b602e09e-97ab-bfbe-1796-bc44ebb7dc28
ms.date: 03/29/2019
localization_priority: Normal
---


# FillFormat object (Excel)

Represents fill formatting for a shape.


## Remarks

A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.

Many of the properties of the **FillFormat** object are read-only. To set one of these properties, you have to apply the corresponding method.


## Example

Use the **[Fill](Excel.Shape.Fill.md)** property of the **Shape** object to return a **FillFormat** object. The following example adds a rectangle to _myDocument_, and then sets the gradient and color for the rectangle's fill.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 90, 90, 90, 80).Fill 
 .ForeColor.RGB = RGB(0, 128, 128) 
 .OneColorGradient msoGradientHorizontal, 1, 1 
End With
```


## Methods

- [OneColorGradient](Excel.FillFormat.OneColorGradient.md)
- [Patterned](Excel.FillFormat.Patterned.md)
- [PresetGradient](Excel.FillFormat.PresetGradient.md)
- [PresetTextured](Excel.FillFormat.PresetTextured.md)
- [Solid](Excel.FillFormat.Solid.md)
- [TwoColorGradient](Excel.FillFormat.TwoColorGradient.md)
- [UserPicture](Excel.FillFormat.UserPicture.md)
- [UserTextured](Excel.FillFormat.UserTextured.md)

## Properties

- [Application](Excel.FillFormat.Application.md)
- [BackColor](Excel.FillFormat.BackColor.md)
- [Creator](Excel.FillFormat.Creator.md)
- [ForeColor](Excel.FillFormat.ForeColor.md)
- [GradientAngle](Excel.FillFormat.GradientAngle.md)
- [GradientColorType](Excel.FillFormat.GradientColorType.md)
- [GradientDegree](Excel.FillFormat.GradientDegree.md)
- [GradientStops](Excel.FillFormat.GradientStops.md)
- [GradientStyle](Excel.FillFormat.GradientStyle.md)
- [GradientVariant](Excel.FillFormat.GradientVariant.md)
- [Parent](Excel.FillFormat.Parent.md)
- [Pattern](Excel.FillFormat.Pattern.md)
- [PictureEffects](Excel.FillFormat.PictureEffects.md)
- [PresetGradientType](Excel.FillFormat.PresetGradientType.md)
- [PresetTexture](Excel.FillFormat.PresetTexture.md)
- [RotateWithObject](Excel.FillFormat.RotateWithObject.md)
- [TextureAlignment](Excel.FillFormat.TextureAlignment.md)
- [TextureHorizontalScale](Excel.FillFormat.TextureHorizontalScale.md)
- [TextureName](Excel.FillFormat.TextureName.md)
- [TextureOffsetX](Excel.FillFormat.TextureOffsetX.md)
- [TextureOffsetY](Excel.FillFormat.TextureOffsetY.md)
- [TextureTile](Excel.FillFormat.TextureTile.md)
- [TextureType](Excel.FillFormat.TextureType.md)
- [TextureVerticalScale](Excel.FillFormat.TextureVerticalScale.md)
- [Transparency](Excel.FillFormat.Transparency.md)
- [Type](Excel.FillFormat.Type.md)
- [Visible](Excel.FillFormat.Visible.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
