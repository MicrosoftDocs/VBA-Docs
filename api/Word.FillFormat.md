---
title: FillFormat object (Word)
keywords: vbawd10.chm2504
f1_keywords:
- vbawd10.chm2504
ms.prod: word
api_name:
- Word.FillFormat
ms.assetid: 39205d07-9e37-1be1-ec4a-93ba8bac2f26
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat object (Word)

Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.


## Remarks

Use the  **Fill** property to return a **FillFormat** object. The following example adds a rectangle to the active document and then sets the gradient and color for the rectangle's fill.


```vb
With ActiveDocument.Shapes _ 
 .AddShape(msoShapeRectangle, 90, 90, 90, 80).Fill 
 .ForeColor.RGB = RGB(0, 128, 128) 
 .OneColorGradient msoGradientHorizontal, 1, 1 
End With
```

Many of the properties of the  **FillFormat** object are read-only. To set one of these properties, you have to apply the corresponding method.


## Methods



|Name|
|:-----|
|[OneColorGradient](Word.FillFormat.OneColorGradient.md)|
|[Patterned](Word.FillFormat.Patterned.md)|
|[PresetGradient](Word.FillFormat.PresetGradient.md)|
|[PresetTextured](Word.FillFormat.PresetTextured.md)|
|[Solid](Word.FillFormat.Solid.md)|
|[TwoColorGradient](Word.FillFormat.TwoColorGradient.md)|
|[UserPicture](Word.FillFormat.UserPicture.md)|
|[UserTextured](Word.FillFormat.UserTextured.md)|

## Properties



|Name|
|:-----|
|[Application](Word.FillFormat.Application.md)|
|[BackColor](Word.FillFormat.BackColor.md)|
|[Creator](Word.FillFormat.Creator.md)|
|[ForeColor](Word.FillFormat.ForeColor.md)|
|[GradientAngle](Word.FillFormat.GradientAngle.md)|
|[GradientColorType](Word.FillFormat.GradientColorType.md)|
|[GradientDegree](Word.FillFormat.GradientDegree.md)|
|[GradientStops](Word.FillFormat.GradientStops.md)|
|[GradientStyle](Word.FillFormat.GradientStyle.md)|
|[GradientVariant](Word.FillFormat.GradientVariant.md)|
|[Parent](Word.FillFormat.Parent.md)|
|[Pattern](Word.FillFormat.Pattern.md)|
|[PictureEffects](Word.FillFormat.PictureEffects.md)|
|[PresetGradientType](Word.FillFormat.PresetGradientType.md)|
|[PresetTexture](Word.FillFormat.PresetTexture.md)|
|[RotateWithObject](Word.FillFormat.RotateWithObject.md)|
|[TextureAlignment](Word.FillFormat.TextureAlignment.md)|
|[TextureHorizontalScale](Word.FillFormat.TextureHorizontalScale.md)|
|[TextureName](Word.FillFormat.TextureName.md)|
|[TextureOffsetX](Word.FillFormat.TextureOffsetX.md)|
|[TextureOffsetY](Word.FillFormat.TextureOffsetY.md)|
|[TextureTile](Word.FillFormat.TextureTile.md)|
|[TextureType](Word.FillFormat.TextureType.md)|
|[TextureVerticalScale](Word.FillFormat.TextureVerticalScale.md)|
|[Transparency](Word.FillFormat.Transparency.md)|
|[Type](Word.FillFormat.Type.md)|
|[Visible](Word.FillFormat.Visible.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]