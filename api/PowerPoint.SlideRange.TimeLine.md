---
title: SlideRange.TimeLine property (PowerPoint)
keywords: vbapp10.chm532035
f1_keywords:
- vbapp10.chm532035
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.TimeLine
ms.assetid: 3d9ad2f6-6d36-dd3d-d564-9bfe97ce08d8
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideRange.TimeLine property (PowerPoint)

Returns a  **[TimeLine](PowerPoint.TimeLine.md)** object that represents the animation timeline for the slide. Read-only.


## Syntax

_expression_. `TimeLine`

_expression_ A variable that represents a [SlideRange](PowerPoint.SlideRange.md) object.


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


[SlideRange Object](PowerPoint.SlideRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]