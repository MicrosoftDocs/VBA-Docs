---
title: ShapeRange.TextEffect property (PowerPoint)
keywords: vbapp10.chm548034
f1_keywords:
- vbapp10.chm548034
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.TextEffect
ms.assetid: 8cf70ead-8534-ef82-5064-21b9929e6f08
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.TextEffect property (PowerPoint)

Returns a **[TextEffectFormat](PowerPoint.TextEffectFormat.md)** object that contains text-effect formatting properties for the specified shape. Read-only.


## Syntax

_expression_.**TextEffect**

_expression_ A variable that represents a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Return value

TextEffectFormat


## Remarks

Applies to  **[ShapeRange](PowerPoint.ShapeRange.md)** objects that represent WordArt.


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


[ShapeRange Object](PowerPoint.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]