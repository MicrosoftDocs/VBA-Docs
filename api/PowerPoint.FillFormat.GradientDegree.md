---
title: FillFormat.GradientDegree property (PowerPoint)
keywords: vbapp10.chm552014
f1_keywords:
- vbapp10.chm552014
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.GradientDegree
ms.assetid: 201380df-f7b4-a38c-e615-2eb490b7042c
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.GradientDegree property (PowerPoint)

Returns a value that indicates how dark or light a one-color gradient fill is. Read-only.


## Syntax

_expression_.**GradientDegree**

_expression_ A variable that represents a **[FillFormat](powerpoint.fillformat.md)** object.


## Return value

Single


## Remarks

A value of 0 (zero) means that black is mixed in with the shape's foreground color to form the gradient; a value of 1 means that white is mixed in; and values between 0 and 1 mean that a darker or lighter shade of the foreground color is mixed in. 

This property is read-only. Use the  **[OneColorGradient](PowerPoint.FillFormat.OneColorGradient.md)** method to set the gradient degree for the fill.


## Example

This example adds a rectangle to _myDocument_ and sets the degree of its fill gradient to match that of the shape named "Rectangle 2." If Rectangle 2 doesn't have a one-color gradient fill, this example fails.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes

    gradDegree1 = .Item("Rectangle 2").Fill.GradientDegree

    With .AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill

        .ForeColor.RGB = RGB(128, 0, 0)

        .OneColorGradient msoGradientHorizontal, 1, gradDegree1

    End With

End With
```


## See also


[FillFormat Object](PowerPoint.FillFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]