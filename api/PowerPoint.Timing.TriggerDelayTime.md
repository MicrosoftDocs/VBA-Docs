---
title: Timing.TriggerDelayTime property (PowerPoint)
keywords: vbapp10.chm653005
f1_keywords:
- vbapp10.chm653005
ms.prod: powerpoint
api_name:
- PowerPoint.Timing.TriggerDelayTime
ms.assetid: 4d14ffb0-e966-4708-ba30-4a9a1fe34766
ms.date: 06/08/2017
localization_priority: Normal
---


# Timing.TriggerDelayTime property (PowerPoint)

Sets or returns the delay, in seconds, from when an animation trigger is enabled. Read/write.


## Syntax

_expression_. `TriggerDelayTime`

_expression_ A variable that represents a [Timing](PowerPoint.Timing.md) object.


## Return value

Single


## Example

The following example adds a shape to a slide, adds an animation to the shape, and instructs the shape to begin the animation three seconds after it is clicked.


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
        .Duration = 5
        .TriggerShape = shpRectangle
        .TriggerType = msoAnimTriggerOnShapeClick
        .TriggerDelayTime = 3
    End With

End Sub
```


## See also


[Timing Object](PowerPoint.Timing.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]