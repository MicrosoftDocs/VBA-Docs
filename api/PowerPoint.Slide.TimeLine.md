---
title: Slide.TimeLine property (PowerPoint)
keywords: vbapp10.chm531031
f1_keywords:
- vbapp10.chm531031
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.TimeLine
ms.assetid: 7dda6e00-5e22-fb2f-91d9-e9c15f8d62bd
ms.date: 06/08/2017
localization_priority: Normal
---


# Slide.TimeLine property (PowerPoint)

Returns a **[TimeLine](PowerPoint.TimeLine.md)** object that represents the animation timeline for the slide. Read-only.


## Syntax

_expression_. `TimeLine`

_expression_ A variable that represents a [Slide](PowerPoint.Slide.md) object.


## Return value

TimeLine


## Example

The following example adds a bouncing animation to the first shape on the first slide.


```vb
Sub NewTimeLineEffect()

    Dim sldFirst As Slide
    Dim shpFirst As Shape

    Set sldFirst = ActivePresentation.Slides(1)
    Set shpFirst = sldFirst.Shapes(1)

    sldFirst.TimeLine.MainSequence.AddEffect _
        Shape:=shpFirst, EffectId:=msoAnimEffectBounce

End Sub
```


## See also


[Slide Object](PowerPoint.Slide.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]