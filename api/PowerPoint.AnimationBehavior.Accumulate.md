---
title: AnimationBehavior.Accumulate property (PowerPoint)
keywords: vbapp10.chm657004
f1_keywords:
- vbapp10.chm657004
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehavior.Accumulate
ms.assetid: 218687c0-6a0e-22ba-a921-efc460986d54
ms.date: 06/08/2017
localization_priority: Normal
---


# AnimationBehavior.Accumulate property (PowerPoint)

Determines whether animation behaviors accumulate. Read/write.


## Syntax

_expression_. `Accumulate`

_expression_ A variable that represents an [AnimationBehavior](PowerPoint.AnimationBehavior.md) object.


## Return value

MsoAnimAccumulate


## Remarks

Use this property in conjunction with the  **[Additive](PowerPoint.AnimationBehavior.Additive.md)** property to combine animation effects.

The value of the  **Accumulate** property can be one of these **MsoAnimAccumulate** constants.



|Constant|Description|
|:-----|:-----|
|**msoAnimAccumulateAlways**| Animation behaviors accumulate.|
|**msoAnimAccumulateNone**| The default. Animation behaviors do not accumulate.|

## Example

The following example allows a specified animation behavior to accumulate with other animation behaviors.


```vb
Sub SetAccumulate()

    Dim animBehavior As AnimationBehavior

    Set animBehavior = ActiveWindow.Selection.SlideRange(1).TimeLine. _
        MainSequence(1).Behaviors(1)

    animBehavior.Accumulate = msoAnimAccumulateAlways

End Sub
```


## See also


[AnimationBehavior Object](PowerPoint.AnimationBehavior.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]