---
title: TimeLine.MainSequence property (PowerPoint)
keywords: vbapp10.chm649003
f1_keywords:
- vbapp10.chm649003
ms.prod: powerpoint
api_name:
- PowerPoint.TimeLine.MainSequence
ms.assetid: b71f83ad-6d92-cc10-9692-a7567ca0a077
ms.date: 06/08/2017
localization_priority: Normal
---


# TimeLine.MainSequence property (PowerPoint)

Returns a  **[Sequence](PowerPoint.Sequence.md)** object that represents the collection of **[Effect](PowerPoint.Effect.md)** objects in the main animation sequence of a slide.


## Syntax

_expression_. `MainSequence`

_expression_ A variable that represents a [TimeLine](PowerPoint.TimeLine.md) object.


## Return value

Sequence


## Remarks

The default value of the  **MainSequence** property is an empty **Sequence** collection. Any attempt to return a value from this property without adding one or more **Effect** objects to the main animation sequence will result in a run-time error.


## Example

The following example adds a boomerang animation to a new shape on a new slide added to the active presentation.


```vb
Sub NewSequence()

    Dim sldNew As Slide
    Dim shpnew As Shape

    Set sldNew = ActivePresentation.Slides.Add(Index:=1, Layout:=ppLayoutBlank)
    Set shpnew = sldNew.Shapes.AddShape(Type:=msoShape5pointStar, _
        Left:=25, Top:=25, Width:=100, Height:=100)

    With sldNew.TimeLine.MainSequence.AddEffect(Shape:=shpnew, _
            EffectId:=msoAnimEffectBoomerang)
        .Timing.Speed = 0.5
        .Timing.Accelerate = 0.2
    End With

End Sub
```


## See also


[TimeLine Object](PowerPoint.TimeLine.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]