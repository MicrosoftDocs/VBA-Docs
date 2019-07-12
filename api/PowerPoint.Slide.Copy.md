---
title: Slide.Copy method (PowerPoint)
keywords: vbapp10.chm531013
f1_keywords:
- vbapp10.chm531013
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.Copy
ms.assetid: 35844287-a2f3-463d-f735-d88f383ad208
ms.date: 06/08/2017
localization_priority: Normal
---


# Slide.Copy method (PowerPoint)

Copies the specified object to the Clipboard.


## Syntax

_expression_.**Copy**

_expression_ A variable that represents a [Slide](PowerPoint.Slide.md) object.


## Remarks

Use the  **Paste** method to paste the contents of the Clipboard.


## Example

This example copies slide one in the active presentation to the Clipboard.


```vb
ActivePresentation.Slides(1).Copy
```


## See also


[Slide Object](PowerPoint.Slide.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]