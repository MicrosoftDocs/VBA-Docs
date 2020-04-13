---
title: EffectParameters.Color2 property (PowerPoint)
keywords: vbapp10.chm654006
f1_keywords:
- vbapp10.chm654006
ms.prod: powerpoint
api_name:
- PowerPoint.EffectParameters.Color2
ms.assetid: 9baff264-9b29-8065-a338-374bdc303451
ms.date: 06/08/2017
localization_priority: Normal
---


# EffectParameters.Color2 property (PowerPoint)

Returns a **[ColorFormat](PowerPoint.ColorFormat.md)** object that represents the color on which to end a color-cycle animation.


## Syntax

_expression_. `Color2`

_expression_ A variable that represents an [EffectParameters](PowerPoint.EffectParameters.md) object.


## Return value

ColorFormat


## Example

The following example adds a shape, adds a fill animation to that shape, then reports the starting and ending fill colors.


```vb
Sub SetStartEndColors()

    Dim effChangeFill As Effect
    Dim shpCube As Shape
    Dim a As AnimationBehavior

    'Adds cube and set fill effect
    Set shpCube = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeCube, Left:=300, _
        Top:=300, Width:=100, Height:=100)

    Set effChangeFill = ActivePresentation.Slides(1).TimeLine _
        .MainSequence.AddEffect(Shape:=shpCube, _
        effectId:=msoAnimEffectChangeFillColor)

    'Sets duration of effect and displays a message containing
    'the starting and ending colors for the fill effect
    effChangeFill.Timing.Duration = 3
    MsgBox "Start Color = " & effChangeFill.EffectParameters _
        .Color1 & vbCrLf & "End Color = " & effChangeFill _
        .EffectParameters.Color2

End Sub
```


## See also



[EffectParameters Object](PowerPoint.EffectParameters.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]