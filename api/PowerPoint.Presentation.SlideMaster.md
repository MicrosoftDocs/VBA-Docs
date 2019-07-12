---
title: Presentation.SlideMaster property (PowerPoint)
keywords: vbapp10.chm583003
f1_keywords:
- vbapp10.chm583003
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.SlideMaster
ms.assetid: 86b11fcd-b979-6ffe-bda7-1b9c6e807d29
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.SlideMaster property (PowerPoint)

Returns a  **[Master](PowerPoint.Master.md)** object that represents the slide master.


## Syntax

_expression_. `SlideMaster`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

Master


## Example

This example sets the background pattern for the slide master for the active presentation.


```vb
Application.ActivePresentation.SlideMaster.Background.Fill _
    .PresetTextured msoTextureGreenMarble
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]