---
title: EffectParameters.FontName property (PowerPoint)
keywords: vbapp10.chm654008
f1_keywords:
- vbapp10.chm654008
ms.prod: powerpoint
api_name:
- PowerPoint.EffectParameters.FontName
ms.assetid: a2f966d5-060e-60b8-422f-b4fab5247736
ms.date: 06/08/2017
localization_priority: Normal
---


# EffectParameters.FontName property (PowerPoint)

Returns or sets the name of the font in the specified WordArt. Read/write.


## Syntax

_expression_. `FontName`

_expression_ A variable that represents an [EffectParameters](PowerPoint.EffectParameters.md) object.


## Return value

String


## Example

This example sets the font name to "Courier New" for shape three on _myDocument_ if the shape is WordArt.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3)

    If .Type = msoTextEffect Then

        .TextEffect.FontName = "Courier New"

    End If

End With
```


## See also


[EffectParameters Object](PowerPoint.EffectParameters.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]