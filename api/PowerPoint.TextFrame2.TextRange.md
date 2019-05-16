---
title: TextFrame2.TextRange property (PowerPoint)
keywords: vbapp10.chm678016
f1_keywords:
- vbapp10.chm678016
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.TextRange
ms.assetid: 288c1209-d12d-fd7c-bc1a-6775d844ca6b
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame2.TextRange property (PowerPoint)

Returns a  **[TextRange2 object (PowerPoint)](PowerPoint.textrange2.md)** object that represents the text in the specified text frame. Read-only.


## Syntax

_expression_. `TextRange2`

 _expression_ An expression that returns a **[TextFrame2](PowerPoint.TextFrame2.md)** object.


## Return value

TextRange2


## Example

This example shows how to set the text for shape one on slide one of the active presentation to the word "Hello!"


```vb
Public Sub TextRange_Example()



    Dim pptSlide As Slide

    Set pptSlide = ActivePresentation.Slides(1)

    pptSlide.Shapes(1).TextFrame2.TextRange = "Hello!"



End Sub
```


## See also


[TextFrame2 Object](PowerPoint.TextFrame2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]