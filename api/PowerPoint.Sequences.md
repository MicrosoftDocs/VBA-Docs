---
title: Sequences object (PowerPoint)
keywords: vbapp10.chm650000
f1_keywords:
- vbapp10.chm650000
ms.prod: powerpoint
api_name:
- PowerPoint.Sequences
ms.assetid: 7650703c-9072-6867-6367-4496b067aa8e
ms.date: 06/08/2017
localization_priority: Normal
---


# Sequences object (PowerPoint)

Represents a collection of  **[Sequence](PowerPoint.Sequence.md)** objects. Use a **Sequence** object to add, find, modify, and clone animation effects.


## Example

Use the [InteractiveSequences](PowerPoint.TimeLine.InteractiveSequences.md)property of the  **[TimeLine](PowerPoint.TimeLine.md)** object to return a **Sequences** collection. Use the [Add](PowerPoint.Sequences.Add.md)method to add an interactive animation sequence. The following example adds two shapes on the first slide of the active presentation and sets interactive effect for the star shape so that when you click the bevel shape, the star shape is be animated.


```vb
Sub AddNewSequence()

    Dim shp1 As Shape
    Dim shp2 As Shape
    Dim interEffect As Effect

    Set shp1 = ActivePresentation.Slides(1).Shapes.AddShape _
        (Type:=msoShape32pointStar, Left:=100, _
        Top:=100, Width:=200, Height:=200)

    Set shp2 = ActivePresentation.Slides(1).Shapes.AddShape _
        (Type:=msoShapeBevel, Left:=400, _
        Top:=200, Width:=150, Height:=100)

    With ActivePresentation.Slides(1).TimeLine.InteractiveSequences.Add(1)
        Set interEffect = .AddEffect(shp2, msoAnimEffectBlinds, _
            trigger:=msoAnimTriggerOnShapeClick)
        interEffect.Shape = shp1
    End With

End Sub
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
