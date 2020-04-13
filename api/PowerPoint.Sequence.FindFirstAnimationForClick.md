---
title: Sequence.FindFirstAnimationForClick method (PowerPoint)
keywords: vbapp10.chm651007
f1_keywords:
- vbapp10.chm651007
ms.prod: powerpoint
api_name:
- PowerPoint.Sequence.FindFirstAnimationForClick
ms.assetid: e1fb9fbe-394e-6765-e4ab-6850e540494b
ms.date: 06/08/2017
localization_priority: Normal
---


# Sequence.FindFirstAnimationForClick method (PowerPoint)

Returns an **[Effect](PowerPoint.Effect.md)** object that represents the first animation started by the specified click number.


## Syntax

_expression_. `FindFirstAnimationForClick`( `_click_` )

_expression_ A variable that represents a [Sequence](PowerPoint.Sequence.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _click_|Required|**Long**|The specified click number.|

## Return value

Effect


## Example

The following example finds the first animation for the first click of the first slide and changes the effect to a bounce.


```vb
Sub FindFirstAnimationClick()

    Dim sldFirst As Slide
    Dim effClick As Effect

    Set sldFirst = ActivePresentation.Slides(1)
    Set effClick = sldFirst.TimeLine.MainSequence _
        .FindFirstAnimationForClick(Click:=1)
    effClick.EffectType = msoAnimEffectBounce

End Sub
```


## See also


[Sequence Object](PowerPoint.Sequence.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]