---
title: SlideRange.PrintSteps property (PowerPoint)
keywords: vbapp10.chm532010
f1_keywords:
- vbapp10.chm532010
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.PrintSteps
ms.assetid: 043a1e60-0810-3f22-7c40-a8a97eb59e4e
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideRange.PrintSteps property (PowerPoint)

Returns the number of slides you'd need to print to simulate the builds on the specified slide, slide master, or range of slides. Read-only.


## Syntax

_expression_. `PrintSteps`

_expression_ A variable that represents a [SlideRange](PowerPoint.SlideRange.md) object.


## Return value

Long


## Example

This example sets a variable to the number of slides you'd need to print to simulate the builds on slide one in the active presentation and then displays the value of the variable.


```vb
steps1 = ActivePresentation.Slides(1).PrintSteps

MsgBox steps1
```


## See also


[SlideRange Object](PowerPoint.SlideRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]