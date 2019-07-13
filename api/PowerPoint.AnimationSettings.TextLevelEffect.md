---
title: AnimationSettings.TextLevelEffect property (PowerPoint)
keywords: vbapp10.chm565011
f1_keywords:
- vbapp10.chm565011
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationSettings.TextLevelEffect
ms.assetid: 008e3db2-2d22-5218-c312-663f0106adc6
ms.date: 06/08/2017
localization_priority: Normal
---


# AnimationSettings.TextLevelEffect property (PowerPoint)

Indicates whether the text in the specified shape is animated by first-level paragraphs, second-level paragraphs, or some other level of paragraphs (up to fifth-level paragraphs). Read/write.


## Syntax

_expression_. `TextLevelEffect`

_expression_ A variable that represents a [AnimationSettings](PowerPoint.AnimationSettings.md) object.


## Return value

PpTextLevelEffect


## Remarks

For the  **TextLevelEffect** property setting to take effect, the **[Animate](PowerPoint.AnimationSettings.Animate.md)** property must be set to **True**.

The value of the  **TextLevelEffect** property can be one of these **PpTextLevelEffect** constants.


||
|:-----|
|**ppAnimateByAllLevels**|
|**ppAnimateByFifthLevel**|
|**ppAnimateByFirstLevel**|
|**ppAnimateByFourthLevel**|
|**ppAnimateBySecondLevel**|
|**ppAnimateByThirdLevel**|
|**ppAnimateLevelMixed**|
|**ppAnimateLevelNone**|

## Example

This example adds a title slide and title text to the active presentation and sets the title to be built letter by letter.


```vb
With ActivePresentation.Slides.Add(1, ppLayoutTitleOnly).Shapes(1)

    .TextFrame.TextRange.Text = "Sample title"

    With .AnimationSettings

        .Animate = True

        .TextLevelEffect = ppAnimateByFirstLevel

        .TextUnitEffect = ppAnimateByCharacter

        .EntryEffect = ppEffectFlyFromLeft

    End With

End With
```


## See also


[AnimationSettings Object](PowerPoint.AnimationSettings.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]