---
title: FillFormat object (PowerPoint)
keywords: vbapp10.chm552000
f1_keywords:
- vbapp10.chm552000
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat
ms.assetid: 5bd4e2cb-4466-b468-d494-bec30ed5c9d8
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat object (PowerPoint)

Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.


## Remarks

Many of the properties of the  **FillFormat** object are read-only. To set one of these properties, you have to apply the corresponding method.


## Example

Use the  **Fill** property to return a **FillFormat** object. The following example adds a rectangle to _myDocument_ and then sets the gradient and color for the rectangle's fill.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes _

        .AddShape(msoShapeRectangle, 90, 90, 90, 80).Fill

    .ForeColor.RGB = RGB(0, 128, 128)

    .OneColorGradient msoGradientHorizontal, 1, 1

End With
```


## Methods



|Name|
|:-----|
|[Background](PowerPoint.FillFormat.Background.md)|
|[OneColorGradient](PowerPoint.FillFormat.OneColorGradient.md)|
|[Patterned](PowerPoint.FillFormat.Patterned.md)|
|[PresetGradient](PowerPoint.FillFormat.PresetGradient.md)|
|[PresetTextured](PowerPoint.FillFormat.PresetTextured.md)|
|[Solid](PowerPoint.FillFormat.Solid.md)|
|[TwoColorGradient](PowerPoint.FillFormat.TwoColorGradient.md)|
|[UserPicture](PowerPoint.FillFormat.UserPicture.md)|
|[UserTextured](PowerPoint.FillFormat.UserTextured.md)|

## Properties



|Name|
|:-----|
|[Application](PowerPoint.FillFormat.Application.md)|
|[BackColor](PowerPoint.FillFormat.BackColor.md)|
|[Creator](PowerPoint.FillFormat.Creator.md)|
|[ForeColor](PowerPoint.FillFormat.ForeColor.md)|
|[GradientAngle](PowerPoint.FillFormat.GradientAngle.md)|
|[GradientColorType](PowerPoint.FillFormat.GradientColorType.md)|
|[GradientDegree](PowerPoint.FillFormat.GradientDegree.md)|
|[GradientStops](PowerPoint.FillFormat.GradientStops.md)|
|[GradientStyle](PowerPoint.FillFormat.GradientStyle.md)|
|[GradientVariant](PowerPoint.FillFormat.GradientVariant.md)|
|[Parent](PowerPoint.FillFormat.Parent.md)|
|[Pattern](PowerPoint.FillFormat.Pattern.md)|
|[PictureEffects](PowerPoint.FillFormat.PictureEffects.md)|
|[PresetGradientType](PowerPoint.FillFormat.PresetGradientType.md)|
|[PresetTexture](PowerPoint.FillFormat.PresetTexture.md)|
|[RotateWithObject](PowerPoint.FillFormat.RotateWithObject.md)|
|[TextureAlignment](PowerPoint.FillFormat.TextureAlignment.md)|
|[TextureHorizontalScale](PowerPoint.FillFormat.TextureHorizontalScale.md)|
|[TextureName](PowerPoint.FillFormat.TextureName.md)|
|[TextureOffsetX](PowerPoint.FillFormat.TextureOffsetX.md)|
|[TextureOffsetY](PowerPoint.FillFormat.TextureOffsetY.md)|
|[TextureTile](PowerPoint.FillFormat.TextureTile.md)|
|[TextureType](PowerPoint.FillFormat.TextureType.md)|
|[TextureVerticalScale](PowerPoint.FillFormat.TextureVerticalScale.md)|
|[Transparency](PowerPoint.FillFormat.Transparency.md)|
|[Type](PowerPoint.FillFormat.Type.md)|
|[Visible](PowerPoint.FillFormat.Visible.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]