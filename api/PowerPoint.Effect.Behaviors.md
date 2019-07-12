---
title: Effect.Behaviors property (PowerPoint)
keywords: vbapp10.chm652017
f1_keywords:
- vbapp10.chm652017
ms.prod: powerpoint
api_name:
- PowerPoint.Effect.Behaviors
ms.assetid: e5335758-2f92-ccbc-a665-b6d5947e79f2
ms.date: 06/08/2017
localization_priority: Normal
---


# Effect.Behaviors property (PowerPoint)

Returns a specified slide animation behavior as an  **[AnimationBehaviors](PowerPoint.AnimationBehaviors.md)** collection.


## Syntax

_expression_. `Behaviors`

_expression_ A variable that represents an [Effect](PowerPoint.Effect.md) object.


## Return value

AnimationBehaviors


## Remarks

To return a single  **[AnimationBehavior](PowerPoint.AnimationBehavior.md)** object in the **AnimationBehaviors** collection, use the **[Item](PowerPoint.AnimationBehaviors.Item.md)** method or **Behaviors** (_index_), where _index_ is the index number of the **AnimationBehavior** object in the **AnimationBehaviors** collection.


## Example

The following example returns a specific animation behavior type in the active presentation.


```vb
Sub ReturnTypeValue
    MsgBox ActiveWindow.Selection.SlideRange(1).TimeLine _
        .MainSequence(1).Behaviors.Item(1).Type
End Sub
```


## See also



[Effect Object](PowerPoint.Effect.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]