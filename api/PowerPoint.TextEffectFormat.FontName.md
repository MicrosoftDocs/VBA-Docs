---
title: TextEffectFormat.FontName property (PowerPoint)
keywords: vbapp10.chm556006
f1_keywords:
- vbapp10.chm556006
ms.prod: powerpoint
api_name:
- PowerPoint.TextEffectFormat.FontName
ms.assetid: 4fdfc7a2-4b2e-e90f-719d-75a3f73c34e6
ms.date: 06/08/2017
localization_priority: Normal
---


# TextEffectFormat.FontName property (PowerPoint)

Returns or sets the name of the font in the specified WordArt. Read/write.


## Syntax

_expression_. `FontName`

_expression_ A variable that represents a [TextEffectFormat](PowerPoint.TextEffectFormat.md) object.


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


[TextEffectFormat Object](PowerPoint.TextEffectFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]