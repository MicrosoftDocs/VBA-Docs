---
title: AnimationBehavior.Timing property (PowerPoint)
keywords: vbapp10.chm657011
f1_keywords:
- vbapp10.chm657011
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehavior.Timing
ms.assetid: 343f11d4-04bf-2637-dbbc-dc3256d57940
ms.date: 06/08/2017
localization_priority: Normal
---


# AnimationBehavior.Timing property (PowerPoint)

Returns a **[Timing](PowerPoint.Timing.md)** object that represents the timing properties for an animation sequence.


## Syntax

_expression_. `Timing`

_expression_ A variable that represents an [AnimationBehavior](PowerPoint.AnimationBehavior.md) object.


## Return value

Timing


## Example

The following example sets the duration of the first animation sequence on the first slide.


```vb
Sub SetTiming()
    ActivePresentation.Slides(1).TimeLine _
        .MainSequence(1).Timing.Duration = 1
End Sub
```


## See also


[AnimationBehavior Object](PowerPoint.AnimationBehavior.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]