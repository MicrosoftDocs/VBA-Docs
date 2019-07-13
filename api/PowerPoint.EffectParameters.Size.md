---
title: EffectParameters.Size property (PowerPoint)
keywords: vbapp10.chm654005
f1_keywords:
- vbapp10.chm654005
ms.prod: powerpoint
api_name:
- PowerPoint.EffectParameters.Size
ms.assetid: cdc1845d-0c6e-36f8-f22e-296aefcc974a
ms.date: 06/08/2017
localization_priority: Normal
---


# EffectParameters.Size property (PowerPoint)

Returns or sets the character size, in points. Read/write.


## Syntax

_expression_.**Size**

_expression_ A variable that represents an [EffectParameters](PowerPoint.EffectParameters.md) object.


## Return value

Single


## Example

This example sets the size of the text attached to shape one on slide one to 24 points.


```vb
Application.ActivePresentation.Slides(1) _
    .Shapes(1).TextFrame.TextRange.Font _
    .Size = 24
```


## See also


[EffectParameters Object](PowerPoint.EffectParameters.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]