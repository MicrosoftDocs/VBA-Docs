---
title: FillFormat.Solid method (PowerPoint)
keywords: vbapp10.chm552007
f1_keywords:
- vbapp10.chm552007
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.Solid
ms.assetid: 0d3302de-2b8b-2a05-697d-0010882588e5
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.Solid method (PowerPoint)

Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.


## Syntax

_expression_.**Solid**

_expression_ A variable that represents a **[FillFormat](powerpoint.fillformat.md)** object.


## Example

This example converts all fills on _myDocument_ to uniform red fills.


```vb
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Shapes

    With s.Fill

        .Solid

        .ForeColor.RGB = RGB(255, 0, 0)

    End With

Next
```


## See also


[FillFormat Object](PowerPoint.FillFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]