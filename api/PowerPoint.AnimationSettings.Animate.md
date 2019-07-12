---
title: AnimationSettings.Animate property (PowerPoint)
keywords: vbapp10.chm565013
f1_keywords:
- vbapp10.chm565013
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationSettings.Animate
ms.assetid: 7434630f-3c73-4261-36f7-a26d45e9df11
ms.date: 06/08/2017
localization_priority: Normal
---


# AnimationSettings.Animate property (PowerPoint)

Determines whether the specified shape is animated during a slide show. Read/write.


## Syntax

_expression_. `Animate`

_expression_ A variable that represents an [AnimationSettings](PowerPoint.AnimationSettings.md) object.


## Return value

MsoTriState


## Remarks

For a shape to be animated, the  **[TextLevelEffect](PowerPoint.AnimationSettings.TextLevelEffect.md)** property of the **AnimationSettings** object for the shape must be set to something other than **ppAnimateLevelNone**, and either the **Animate** property must be set to **True**, or the **[EntryEffect](PowerPoint.AnimationSettings.EntryEffect.md)** property must be set to a constant other than **ppEffectNone**.

The value of the  **Animate** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The specified shape is not animated during a slide show.|
|**msoTrue**| The specified shape is animated during a slide show.|

## Example

This example specifies that the title on slide two in the active presentation appear dimmed after the title is built. If the title is the last or only shape to be built on slide two, the text won't be dimmed.


```vb
With ActivePresentation.Slides(2).Shapes.Title.AnimationSettings

    .TextLevelEffect = ppAnimateByAllLevels

    .AfterEffect = ppAfterEffectDim

    .Animate = msoTrue

End With
```


## See also


[AnimationSettings Object](PowerPoint.AnimationSettings.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]