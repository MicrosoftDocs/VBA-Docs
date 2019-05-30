---
title: FillFormat object (Publisher)
keywords: vbapb10.chm2424831
f1_keywords:
- vbapb10.chm2424831
ms.prod: publisher
api_name:
- Publisher.FillFormat
ms.assetid: 0a5d4f7a-c42a-28ad-c86d-ac9828a3b874
ms.date: 05/31/2019
localization_priority: Normal
---


# FillFormat object (Publisher)

Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semitransparent fill.
 

## Remarks

Many of the properties of the **FillFormat** object are read-only. To set one of these properties, you have to apply the corresponding method.
 
Use the **[Shape.Fill](Publisher.Shape.Fill.md)** property to return a **FillFormat** object. 
 

## Example

The following example adds a shape to the active document and then sets the gradient and color for the shape's fill.

```vb
Sub AddShapeAndSetFill() 
 With ActiveDocument.Pages(1).Shapes.AddShape(Type:=msoShapeHeart, _ 
 Left:=90, Top:=90, Width:=90, Height:=80).Fill 
 .ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 .OneColorGradient Style:=msoGradientHorizontal, _ 
 Variant:=1, Degree:=1 
 End With 
End Sub
```


## Methods

- [OneColorGradient](Publisher.FillFormat.OneColorGradient.md)
- [Patterned](Publisher.FillFormat.Patterned.md)
- [PresetGradient](Publisher.FillFormat.PresetGradient.md)
- [PresetTextured](Publisher.FillFormat.PresetTextured.md)
- [Solid](Publisher.FillFormat.Solid.md)
- [TwoColorGradient](Publisher.FillFormat.TwoColorGradient.md)
- [UserPicture](Publisher.FillFormat.UserPicture.md)
- [UserTextured](Publisher.FillFormat.UserTextured.md)

## Properties

- [Application](Publisher.FillFormat.Application.md)
- [BackColor](Publisher.FillFormat.BackColor.md)
- [ForeColor](Publisher.FillFormat.ForeColor.md)
- [GradientAngle](Publisher.fillformat.gradientangle.md)
- [GradientColorType](Publisher.FillFormat.GradientColorType.md)
- [GradientDegree](Publisher.FillFormat.GradientDegree.md)
- [GradientStyle](Publisher.FillFormat.GradientStyle.md)
- [GradientVariant](Publisher.FillFormat.GradientVariant.md)
- [Parent](Publisher.FillFormat.Parent.md)
- [Pattern](Publisher.FillFormat.Pattern.md)
- [PresetGradientType](Publisher.FillFormat.PresetGradientType.md)
- [PresetTexture](Publisher.FillFormat.PresetTexture.md)
- [RotateWithObject](Publisher.fillformat.rotatewithobject.md)
- [TextureAlignment](Publisher.fillformat.texturealignment.md)
- [TextureHorizontalScale](Publisher.fillformat.texturehorizontalscale.md)
- [TextureName](Publisher.FillFormat.TextureName.md)
- [TextureOffsetX](Publisher.fillformat.textureoffsetx.md)
- [TextureOffsetY](Publisher.fillformat.textureoffsety.md)
- [TextureType](Publisher.FillFormat.TextureType.md)
- [TextureVerticalScale](Publisher.fillformat.textureverticalscale.md)
- [Transparency](Publisher.fillformat.transparency.md)
- [Type](Publisher.FillFormat.Type.md)
- [Visible](Publisher.FillFormat.Visible.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]