---
title: Effect object (PowerPoint)
keywords: vbapp10.chm652000
f1_keywords:
- vbapp10.chm652000
ms.prod: powerpoint
api_name:
- PowerPoint.Effect
ms.assetid: 359ac3da-86cd-8003-d691-349d20fd1777
ms.date: 06/08/2017
localization_priority: Normal
---


# Effect object (PowerPoint)

Represents timing information about a slide animation.


## Example

Use the [AddEffect](PowerPoint.Sequence.AddEffect.md)method to add an effect. This example adds a shape to the first slide in the active presentation and adds an effect and a behavior to the shape.


```vb
Sub NewShapeAndEffect()

    Dim shpStar As Shape

    Dim sldOne As Slide

    Dim effNew As Effect



    Set sldOne = ActivePresentation.Slides(1)

    Set shpStar = sldOne.Shapes.AddShape(Type:=msoShape5pointStar, _

        Left:=150, Top:=72, Width:=400, Height:=400)

    Set effNew = sldOne.TimeLine.MainSequence.AddEffect(Shape:=shpStar, _

        EffectId:=msoAnimEffectStretchy, Trigger:=msoAnimTriggerAfterPrevious)

    With effNew

        With .Behaviors.Add(msoAnimTypeScale).ScaleEffect

            .FromX = 75

            .FromY = 75

            .ToX = 0

            .ToY = 0

        End With

        .Timing.AutoReverse = msoTrue

    End With

End Sub
```

To refer to an existing  **Effect** object, use **[MainSequence](PowerPoint.TimeLine.MainSequence.md)** (_index_), where _index_ is the number of the **Effect** object in the **[Sequence](PowerPoint.Sequence.md)** collection. This example changes the effect for the first sequence and specifies the behavior for that effect.




```vb
Sub ChangeEffect()

    With ActivePresentation.Slides(1).TimeLine _

        .MainSequence(1)

        .EffectType = msoAnimEffectSpin

        With .Behaviors(1).RotationEffect

            .From = 100

            .To = 360

            .By = 5

        End With

    End With

End Sub
```


## Methods



|Name|
|:-----|
|**[Delete](PowerPoint.Effect.Delete.md)**|
|**[MoveAfter](PowerPoint.Effect.MoveAfter.md)**|
|**[MoveBefore](PowerPoint.Effect.MoveBefore.md)**|
|**[MoveTo](PowerPoint.Effect.MoveTo.md)**|

## Properties



|Name|
|:-----|
|**[Application](PowerPoint.Effect.Application.md)**|
|**[Behaviors](PowerPoint.Effect.Behaviors.md)**|
|**[DisplayName](PowerPoint.Effect.DisplayName.md)**|
|**[EffectInformation](PowerPoint.Effect.EffectInformation.md)**|
|**[EffectParameters](PowerPoint.Effect.EffectParameters.md)**|
|**[EffectType](PowerPoint.Effect.EffectType.md)**|
|**[Exit](PowerPoint.Effect.Exit.md)**|
|**[Index](PowerPoint.Effect.Index.md)**|
|**[Paragraph](PowerPoint.Effect.Paragraph.md)**|
|**[Parent](PowerPoint.Effect.Parent.md)**|
|**[Shape](PowerPoint.Effect.Shape.md)**|
|**[TextRangeLength](PowerPoint.Effect.TextRangeLength.md)**|
|**[TextRangeStart](PowerPoint.Effect.TextRangeStart.md)**|
|**[Timing](PowerPoint.Effect.Timing.md)**|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]