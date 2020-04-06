---
title: Shape.TextEffect property (PowerPoint)
keywords: vbapp10.chm547034
f1_keywords:
- vbapp10.chm547034
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.TextEffect
ms.assetid: b5d0a0a5-462d-1ede-3dac-7bedaaa1e318
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.TextEffect property (PowerPoint)

Returns a **[TextEffectFormat](PowerPoint.TextEffectFormat.md)** object that contains text-effect formatting properties for the specified shape. Read-only.


## Syntax

_expression_.**TextEffect**

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Return value

TextEffectFormat


## Remarks

Applies to  **[Shape](PowerPoint.Shape.md)** objects that represent WordArt.


## Example

This example sets the font style to bold for shape three on _myDocument_ if the shape is WordArt.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3)

    If .Type = msoTextEffect Then

        .TextEffect.FontBold = True

    End If

End With
```


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]