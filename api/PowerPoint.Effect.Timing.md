---
title: Effect.Timing property (PowerPoint)
keywords: vbapp10.chm652009
f1_keywords:
- vbapp10.chm652009
ms.prod: powerpoint
api_name:
- PowerPoint.Effect.Timing
ms.assetid: 88b4f9a5-62aa-6844-e784-f74a1d78aa82
ms.date: 06/08/2017
localization_priority: Normal
---


# Effect.Timing property (PowerPoint)

Returns a **[Timing](PowerPoint.Timing.md)** object that represents the timing properties for an animation sequence.


## Syntax

_expression_. `Timing`

_expression_ A variable that represents an [Effect](PowerPoint.Effect.md) object.


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



[Effect Object](PowerPoint.Effect.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]