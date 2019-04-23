---
title: AnimationPoint object (PowerPoint)
keywords: vbapp10.chm664000
f1_keywords:
- vbapp10.chm664000
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationPoint
ms.assetid: 79aa1a47-abab-f98f-955a-48be10a94c41
ms.date: 06/08/2017
localization_priority: Normal
---


# AnimationPoint object (PowerPoint)

Represents an individual animation point for an animation behavior. The  **AnimationPoint** object is a member of the **[AnimationPoints](PowerPoint.AnimationPoints.md)** collection. The **AnimationPoints** collection contains all the animation points for an animation behavior.


## Example

To add or reference an  **AnimationPoint** object, use the [Add](PowerPoint.AnimationPoints.Add.md) or [Item](PowerPoint.AnimationPoints.Item.md)method, respectively. Use the [Time](PowerPoint.AnimationPoint.Time.md)property of an  **AnimationPoint** object to set timing between animation points. Use the **[Value](PowerPoint.AnimationPoint.Value.md)** property to set other animation point properties, such as color. The following example adds three animation points to the first behavior in the active presentation's main animation sequence, and then it changes colors at each animation point.


```vb
Sub AniPoint()

    Dim sldNewSlide As Slide
    Dim shpHeart As Shape
    Dim effCustom As Effect
    Dim aniBehavior As AnimationBehavior
    Dim aptNewPoint As AnimationPoint

    Set sldNewSlide = ActivePresentation.Slides.Add _
        (Index:=1, Layout:=ppLayoutBlank)

    Set shpHeart = sldNewSlide.Shapes.AddShape _
        (Type:=msoShapeHeart, Left:=100, Top:=100, _
        Width:=200, Height:=200)

    Set effCustom = sldNewSlide.TimeLine.MainSequence _
        .AddEffect(shpHeart, msoAnimEffectCustom)

    Set aniBehavior = effCustom.Behaviors.Add(msoAnimTypeProperty)

    With aniBehavior.PropertyEffect
        .Property = msoAnimShapeFillColor
        Set aptNewPoint = .Points.Add
        aptNewPoint.Time = 0.2
        aptNewPoint.Value = RGB(0, 0, 0)
        Set aptNewPoint = .Points.Add
        aptNewPoint.Time = 0.5
        aptNewPoint.Value = RGB(0, 255, 0)
        Set aptNewPoint = .Points.Add
        aptNewPoint.Time = 1
        aptNewPoint.Value = RGB(0, 255, 255)
    End With

End Sub
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]