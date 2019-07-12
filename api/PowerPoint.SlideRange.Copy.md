---
title: SlideRange.Copy method (PowerPoint)
keywords: vbapp10.chm532013
f1_keywords:
- vbapp10.chm532013
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.Copy
ms.assetid: d781370d-8107-efaa-77ea-a7f1aa58737b
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideRange.Copy method (PowerPoint)

Copies the specified object to the Clipboard.


## Syntax

_expression_.**Copy**

_expression_ A variable that represents a [SlideRange](PowerPoint.SlideRange.md) object.


## Remarks

Use the  **Paste** method to paste the contents of the Clipboard.


## Example

This example copies slide one in the active presentation to the Clipboard.


```vb
ActivePresentation.Slides(1).Copy
```


## See also


[SlideRange Object](PowerPoint.SlideRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]