---
title: AnimationBehavior.MotionEffect property (PowerPoint)
keywords: vbapp10.chm657006
f1_keywords:
- vbapp10.chm657006
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehavior.MotionEffect
ms.assetid: ef9601ab-7a01-ba03-a5ef-a50c4d2c3c79
ms.date: 06/08/2017
localization_priority: Normal
---


# AnimationBehavior.MotionEffect property (PowerPoint)

Returns a **[MotionEffect](PowerPoint.MotionEffect.md)** object that represents the properties of a motion animation.


## Syntax

_expression_. `MotionEffect`

_expression_ A variable that represents an [AnimationBehavior](PowerPoint.AnimationBehavior.md) object.


## Return value

MotionEffect


## Example

This example adds a new motion behavior to the first slide's main sequence that moves the specified animation sequence from one side of the page to the shape's original position.


```vb
Sub NewMotion()

    With ActivePresentation.Slides(1).TimeLine.MainSequence(1) _
            .Behaviors.Add(msoAnimTypeMotion).MotionEffect
        .FromX = 100
        .FromY = 100
        .ToX = 0
        .ToY = 0
    End With

End Sub
```


## See also


[AnimationBehavior Object](PowerPoint.AnimationBehavior.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]