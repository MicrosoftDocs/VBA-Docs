---
title: Slide.PrintSteps property (PowerPoint)
keywords: vbapp10.chm531010
f1_keywords:
- vbapp10.chm531010
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.PrintSteps
ms.assetid: b5474b85-0c1f-aa18-da9d-be7d778e9e16
ms.date: 06/08/2017
localization_priority: Normal
---


# Slide.PrintSteps property (PowerPoint)

Returns the number of slides you'd need to print to simulate the builds on the specified slide, slide master, or range of slides. Read-only.


## Syntax

_expression_. `PrintSteps`

_expression_ A variable that represents a [Slide](PowerPoint.Slide.md) object.


## Return value

Long


## Example

This example sets a variable to the number of slides you'd need to print to simulate the builds on slide one in the active presentation and then displays the value of the variable.


```vb
steps1 = ActivePresentation.Slides(1).PrintSteps

MsgBox steps1
```


## See also


[Slide Object](PowerPoint.Slide.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]