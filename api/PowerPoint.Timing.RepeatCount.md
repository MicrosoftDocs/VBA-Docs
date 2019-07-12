---
title: Timing.RepeatCount property (PowerPoint)
keywords: vbapp10.chm653007
f1_keywords:
- vbapp10.chm653007
ms.prod: powerpoint
api_name:
- PowerPoint.Timing.RepeatCount
ms.assetid: 71d31607-6006-f2c0-cfa3-3711791331bc
ms.date: 06/08/2017
localization_priority: Normal
---


# Timing.RepeatCount property (PowerPoint)

Sets or returns the number of times to repeat an animation. Read/write.


## Syntax

_expression_. `RepeatCount`

_expression_ A variable that represents a [Timing](PowerPoint.Timing.md) object.


## Return value

Long


## Example

This example creates a shape and adds an animation to it, then repeats the animation twice.


```vb
Sub AddShapeSetTiming()

    Dim effDiamond As Effect
    Dim shpRectangle As Shape

    Set shpRectangle = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeRectangle, Left:=100, _
        Top:=100, Width:=50, Height:=50)

    Set effDiamond = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpRectangle, effectId:=msoAnimEffectPathDiamond)

    With effDiamond.Timing
        .Duration = 5 ' Length of effect.
        .RepeatCount = 2 ' How many times to repeat.
    End With

End Sub
```


## See also


[Timing Object](PowerPoint.Timing.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]