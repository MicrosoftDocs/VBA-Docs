---
title: AnimationPoints object (PowerPoint)
keywords: vbapp10.chm663000
f1_keywords:
- vbapp10.chm663000
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationPoints
ms.assetid: 6ea9ebc4-791c-9781-38c3-8b0973e0d152
ms.date: 06/08/2017
localization_priority: Normal
---


# AnimationPoints object (PowerPoint)

Represents a collection of animation points for a  **[PropertyEffect](PowerPoint.PropertyEffect.md)** object.


## Example

Use the [Points](PowerPoint.PropertyEffect.Points.md)property of the  **[PropertyEffect](PowerPoint.PropertyEffect.md)** object to return an **AnimationPoints** collection object. The following example adds an animation point to the first behavior in the active presentation's main animation sequence.


```vb
Sub AddPoint()
    ActivePresentation.Slides(1).TimeLine.MainSequence(1) _
        .Behaviors(1).PropertyEffect.Points.Add
End Sub
```

Transitions from one animation point to another can sometimes be abrupt or choppy. Use the [Smooth](PowerPoint.AnimationPoints.Smooth.md)property to make transitions smoother. This example smooths the transitions between animation points.




```vb
Sub SmoothTransition()
    ActivePresentation.Slides(1).TimeLine.MainSequence(1) _
        .Behaviors(1).PropertyEffect.Points.Smooth = msoTrue
End Sub
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]