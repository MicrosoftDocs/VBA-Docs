---
title: SlideShowTransition.AdvanceOnTime property (PowerPoint)
keywords: vbapp10.chm539004
f1_keywords:
- vbapp10.chm539004
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowTransition.AdvanceOnTime
ms.assetid: 934c5acc-b230-2b7b-f0f2-4647cce5b62d
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowTransition.AdvanceOnTime property (PowerPoint)

Determines whether the specified slide advances automatically after a specified amount of time has elapsed. Read/write.


## Syntax

_expression_. `AdvanceOnTime`

_expression_ A variable that represents an [SlideShowTransition](PowerPoint.SlideShowTransition.md) object.


## Return value

MsoTriState


## Remarks

Use the  **[AdvanceTime](PowerPoint.SlideShowTransition.AdvanceTime.md)** property to specify the number of seconds after which the slide automatically advances. Set the **[AdvanceMode](PowerPoint.SlideShowSettings.AdvanceMode.md)** property of the **SlideShowSettings** object to **ppSlideShowUseSlideTimings** to put the slide interval settings into effect for the entire slide show.

The value of the  **AdvanceOnTime** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The specified slide does not advance automatically after a specified amount of time has elapsed. |
|**msoTrue**| The specified slide advances automatically after a specified amount of time has elapsed.|

## Example

This example sets slide one in the active presentation to advance after five seconds have passed or when the mouse is clicked&mdash;whichever occurs first.


```vb
With ActivePresentation.Slides(1).SlideShowTransition

    .AdvanceOnClick = msoTrue

    .AdvanceOnTime = msoTrue

    .AdvanceTime = 5

End With
```


## See also


[SlideShowTransition Object](PowerPoint.SlideShowTransition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]