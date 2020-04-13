---
title: RotationEffect object (PowerPoint)
keywords: vbapp10.chm661000
f1_keywords:
- vbapp10.chm661000
ms.prod: powerpoint
api_name:
- PowerPoint.RotationEffect
ms.assetid: d0fc5520-dbbd-a44a-b811-51fd299c4587
ms.date: 06/08/2017
localization_priority: Normal
---


# RotationEffect object (PowerPoint)

Represents a rotation effect for an **[AnimationBehavior](PowerPoint.AnimationBehavior.md)** object.


## Example

Use the [RotationEffect](PowerPoint.AnimationBehavior.RotationEffect.md)property of the  **AnimationBehavior** object to return a **RotationEffect** object. The following example refers to the rotation effect for a given animation behavior.


```vb
ActivePresentation.Slides(1).TimeLine.MainSequence.Item.Behaviors(1).RotationEffect
```

Use the [By](PowerPoint.RotationEffect.By.md), [From](PowerPoint.RotationEffect.From.md), and [To](PowerPoint.RotationEffect.To.md)properties of the  **RotationEffect** object to affect an object's animation rotation. The following example adds a new shape to the first slide and sets the rotation animation behavior.




```vb
Sub AddRotation()

    Dim shpNew As Shape
    Dim effNew As Effect
    Dim aniNew As AnimationBehavior

    Set shpNew = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShape5pointStar, Left:=0, _
        Top:=0, Width:=100, Height:=100)

    Set effNew = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpNew, effectId:=msoAnimEffectCustom)
    Set aniNew = effNew.Behaviors.Add(msoAnimTypeRotation)

    With aniNew.RotationEffect
        'Rotate 270 degrees from current position
        .By = 270
    End With

End Sub
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]