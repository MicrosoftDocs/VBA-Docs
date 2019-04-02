---
title: AnimationBehaviors object (PowerPoint)
keywords: vbapp10.chm656000
f1_keywords:
- vbapp10.chm656000
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehaviors
ms.assetid: 40e11093-5cbd-c8d3-04b5-4cd7de97bfa7
ms.date: 06/08/2017
localization_priority: Normal
---


# AnimationBehaviors object (PowerPoint)

Represents a collection of  **[AnimationBehavior](PowerPoint.AnimationBehavior.md)** objects.


## Example

Use the [Add](PowerPoint.AnimationBehaviors.Add.md)method to add an animation behavior. The following example adds a five-second animated rotation behavior to the main animation sequence on the first slide.


```vb
Sub AnimationObject()

    Dim timeMain As TimeLine



    'Reference the main animation timeline

    Set timeMain = ActivePresentation.Slides(1).TimeLine



    'Add a five-second animated rotation behavior

    'as the first animation in the main animation sequence

    timeMain.MainSequence(1).Behaviors.Add Type:=msoAnimTypeRotation, Index:=1

End Sub
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]