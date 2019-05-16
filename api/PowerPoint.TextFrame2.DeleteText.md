---
title: TextFrame2.DeleteText method (PowerPoint)
keywords: vbapp10.chm678019
f1_keywords:
- vbapp10.chm678019
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.DeleteText
ms.assetid: 47197c75-99be-4f42-0b4a-bf9207480a94
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame2.DeleteText method (PowerPoint)

Deletes the text from a text frame and all the associated properties of the text, including font attributes.


## Syntax

_expression_.**DeleteText**

 _expression_ An expression that returns a **[TextFrame2](PowerPoint.TextFrame2.md)** object.


## Return value

Nothing


## Example

This example shows how to delete the text from shape one on slide one of the active presentation, if that shape contains text.


```vb
Public Sub DeleteText_Example()



    Dim pptSlide As Slide

    Set pptSlide = ActivePresentation.Slides(1)

    pptSlide.Shapes(1).TextFrame2.DeleteText



End Sub
```


## See also


[TextFrame2 Object](PowerPoint.TextFrame2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]