---
title: AnimationSettings.AnimateTextInReverse property (PowerPoint)
keywords: vbapp10.chm565015
f1_keywords:
- vbapp10.chm565015
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationSettings.AnimateTextInReverse
ms.assetid: cceba8ad-9896-10ef-5c11-7c93d370c82c
ms.date: 06/08/2017
localization_priority: Normal
---


# AnimationSettings.AnimateTextInReverse property (PowerPoint)

Determines whether the specified shape is built in reverse order. Applies only to shapes (such as shapes containing lists) that can be built in more than one step. Read/write.


## Syntax

_expression_. `AnimateTextInReverse`

_expression_ A variable that represents an [AnimationSettings](PowerPoint.AnimationSettings.md) object.


## Return value

MsoTriState


## Remarks

The value of the  **AnimateTextInReverse Property** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The specified shape is not built in reverse order.|
|**msoTrue**| The specified shape is built in reverse order.|

You do not see the effects of setting this property unless the specified shape gets animated. For a shape to be animated, the  **TextLevelEffect** property of the **AnimationSettings** object for the shape must be set to something other than **ppAnimateLevelNone** and the **[Animate](PowerPoint.AnimationSettings.Animate.md)** property must be set to **True**.


## Example

This example adds a slide after slide one in the active presentation, sets the title text, adds a three-item list to the text placeholder, and sets the list to be built in reverse order.


```vb
With ActivePresentation.Slides.Add(2, ppLayoutText).Shapes
    .Item(1).TextFrame.TextRange.Text = "Top Three Reasons"
    With .Item(2)
        .TextFrame.TextRange = "Reason 1" & Chr(13) _
            & "Reason 2" & Chr(13) & "Reason 3"
        With .AnimationSettings
            .Animate = msoTrue
            .TextLevelEffect = ppAnimateByFirstLevel
            .AnimateTextInReverse = msoTrue
        End With
    End With
End With
```


## See also


[AnimationSettings Object](PowerPoint.AnimationSettings.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]