---
title: Application.SlideShowEnd event (PowerPoint)
keywords: vbapp10.chm621014
f1_keywords:
- vbapp10.chm621014
ms.prod: powerpoint
api_name:
- PowerPoint.Application.SlideShowEnd
ms.assetid: e46f8177-e00b-6704-1606-dbf9e96bf812
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SlideShowEnd event (PowerPoint)

Occurs after a slide show ends, immediately after the last  **[SlideShowNextSlide](PowerPoint.Application.SlideShowNextSlide.md)** event occurs.


## Syntax

_expression_. `SlideShowEnd`( `_Pres_` )

 _expression_ An expression that returns an **[Application](PowerPoint.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Pres_|Required|**Presentation**|The presentation closed when this event occurs.|

## Remarks

The **SlideShowEnd** event always occurs before a slide show ends if the **[SlideShowBegin](PowerPoint.Application.SlideShowBegin.md)** event has occurred. You can use the **SlideShowEnd** event to return any property settings and variable initializations that occur in the **SlideShowBegin** event to their original settings.

For information about using events with the  **Application** object, see [How to: Use Events with the Application Object](../powerpoint/How-to/use-events-with-the-application-object.md).


## Example

This example turns off the entry effect and automatic advance timing slide show transition effects for slides one through four at the end of the slide show. It also sets the slides to advance manually.


```vb
Private Sub App_SlideShowEnd(ByVal Pres As Presentation)

    With Pres.Slides.Range(Array(1, 4)) _
            .SlideShowTransition

        .EntryEffect = ppEffectNone

        .AdvanceOnTime = msoFalse

    End With



    With Pres.SlideShowSettings

        .AdvanceMode = ppSlideShowManualAdvance

    End With

End Sub
```


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]