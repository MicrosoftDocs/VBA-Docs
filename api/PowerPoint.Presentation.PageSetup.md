---
title: Presentation.PageSetup property (PowerPoint)
keywords: vbapp10.chm583012
f1_keywords:
- vbapp10.chm583012
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.PageSetup
ms.assetid: 81327801-ad21-967c-9682-54a847f79e29
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.PageSetup property (PowerPoint)

Returns a **[PageSetup](PowerPoint.PageSetup.md)** object whose properties control slide setup attributes for the specified presentation. Read-only.


## Syntax

_expression_.**PageSetup**

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

PageSetup


## Example

The following example sets the slide size and slide orientation for the presentation named "Pres1."


```vb
With Presentations("pres1").PageSetup

    .SlideSize = ppSlideSize35MM

    .SlideOrientation = msoOrientationHorizontal

End With
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]