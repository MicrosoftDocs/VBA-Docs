---
title: Presentation.HandoutMaster property (PowerPoint)
keywords: vbapp10.chm583010
f1_keywords:
- vbapp10.chm583010
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.HandoutMaster
ms.assetid: d80a8e51-61db-8da0-1fda-20a043e62569
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.HandoutMaster property (PowerPoint)

Returns a **[Master](PowerPoint.Master.md)** object that represents the handout master. Read-only.


## Syntax

_expression_. `HandoutMaster`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

Master


## Example

This example sets the background pattern on the handout master in the active presentation.


```vb
Application.ActivePresentation.HandoutMaster.Background.Fill _
    .Patterned msoPatternDarkHorizontal
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]