---
title: MotionEffect.FromY property (PowerPoint)
keywords: vbapp10.chm658006
f1_keywords:
- vbapp10.chm658006
ms.prod: powerpoint
api_name:
- PowerPoint.MotionEffect.FromY
ms.assetid: e9754587-f8bc-4af5-bd2b-ec6b1e3c4bc4
ms.date: 06/08/2017
localization_priority: Normal
---


# MotionEffect.FromY property (PowerPoint)

Returns or sets a **Single** that represents the vertical position of a **[MotionEffect](PowerPoint.MotionEffect.md)** object, specified as a percentage of the screen width. Read/write.


## Syntax

_expression_. `FromY`

_expression_ A variable that represents a [MotionEffect](PowerPoint.MotionEffect.md) object.


## Return value

Single


## Remarks

The default value of this property is **Empty**, in which case the current position of the object is used.

Use this property in conjunction with the **ToY** property to resize or jump from one position to another.

Do not confuse this property with the **From** property of the **[ColorEffect](PowerPoint.ColorEffect.md)**, **[RotationEffect](PowerPoint.RotationEffect.md)**, or **[PropertyEffect](PowerPoint.PropertyEffect.md)** objects, which is used to set or change colors, rotations, or other properties of an animation behavior, respectively.


## Example

The following example adds an animation path and sets the starting and ending horizontal and vertical positions.


```vb
Sub AddMotionPath()

    Dim effCustom As Effect
    Dim animMotion As AnimationBehavior
    Dim shpRectangle As Shape

    'Adds shape and sets effect and animation properties
    Set shpRectangle = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeRectangle, Left:=100, _
        Top:=100, Width:=50, Height:=50)

    Set effCustom = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpRectangle, effectId:=msoAnimEffectCustom)

    Set animMotion = effCustom.Behaviors.Add(msoAnimTypeMotion)

    'Sets starting and ending horizontal and vertical positions
    With animMotion.MotionEffect
        .FromX = 0
        .FromY = 0
        .ToX = 50
        .ToY = 50
    End With

End Sub
```


## See also


[MotionEffect Object](PowerPoint.MotionEffect.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]