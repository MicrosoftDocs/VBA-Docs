---
title: SlideRange.Design property (PowerPoint)
keywords: vbapp10.chm532033
f1_keywords:
- vbapp10.chm532033
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.Design
ms.assetid: 7960f99a-fa5a-1ba0-e39a-fe3afe579621
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideRange.Design property (PowerPoint)

Returns a **Design** object representing a design.


## Syntax

_expression_. `Design`

_expression_ A variable that represents a [SlideRange](PowerPoint.SlideRange.md) object.


## Return value

Design


## Example

The following example adds a title master.


```vb
Sub AddDesignMaster

    ActivePresentation.Slides(1).Design.AddTitleMaster

End Sub
```


## See also


[SlideRange Object](PowerPoint.SlideRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]