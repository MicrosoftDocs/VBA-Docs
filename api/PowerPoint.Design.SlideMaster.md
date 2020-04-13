---
title: Design.SlideMaster property (PowerPoint)
keywords: vbapp10.chm644003
f1_keywords:
- vbapp10.chm644003
ms.prod: powerpoint
api_name:
- PowerPoint.Design.SlideMaster
ms.assetid: c6a9263c-462a-e9d8-7afc-32da3e133e90
ms.date: 06/08/2017
localization_priority: Normal
---


# Design.SlideMaster property (PowerPoint)

Returns a **[Master](PowerPoint.Master.md)** object that represents the slide master. Read-only.


## Syntax

_expression_. `SlideMaster`

_expression_ A variable that represents a [Design](PowerPoint.Design.md) object.


## Return value

Master


## Example

This example sets the background pattern for the slide master for the active presentation.


```vb
Application.ActivePresentation.SlideMaster.Background.Fill _
    .PresetTextured msoTextureGreenMarble
```


## See also


[Design Object](PowerPoint.Design.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]