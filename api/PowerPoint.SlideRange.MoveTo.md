---
title: SlideRange.MoveTo Method (PowerPoint)
keywords: vbapp10.chm532034
f1_keywords:
- vbapp10.chm532034
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.MoveTo
ms.assetid: 42be6065-244f-ba03-1f96-5f953258bdb9
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideRange.MoveTo Method (PowerPoint)

Moves the specified object to a specific location within the same collection, renumbering all other items in the collection appropriately.


## Syntax

 _expression_. `MoveTo`( `_toPos_` )

 _expression_ A variable that represents a [SlideRange](./PowerPoint.SlideRange.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _toPos_|Required|**Long**|The index position to which to move the animation effect.|

## Example

This example moves an animation effect to the second in the animation effects collection for the specified shape.


```vb
Sub MoveEffect()

    Dim sldFirst as Slide
    Dim shpFirst As Shape
    Dim effAdd As Effect

    Set sldFirst = ActivePresentation.Slides(1)
    Set shpFirst = sldFirst.Shapes(1)

    Set effAdd = sldFirst.TimeLine.MainSequence.AddEffect _
        (Shape:=shpFirst, effectId:=msoAnimEffectBlinds)

    effAdd.MoveTo toPos:=2

End Sub
```

This example moves the second slide in the active presentation to the first slide.




```vb
Sub MoveSlideToNewLocation()

    ActivePresentation.Slides(2).MoveTo toPos:=1

End Sub
```


## See also


[SlideRange Object](PowerPoint.SlideRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]