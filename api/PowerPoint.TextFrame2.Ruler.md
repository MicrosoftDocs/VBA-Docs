---
title: TextFrame2.Ruler property (PowerPoint)
keywords: vbapp10.chm678018
f1_keywords:
- vbapp10.chm678018
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.Ruler
ms.assetid: 2fcf6db9-e34f-0dac-de6f-3b470d325ee0
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame2.Ruler property (PowerPoint)

Returns a **Ruler2** object that represents the ruler for the specified text. Read-only.


## Syntax

_expression_. `Ruler`

 _expression_ An expression that returns a **[TextFrame2](PowerPoint.TextFrame2.md)** object.


## Return value

Ruler2


## Example

This example shows how to set a left-aligned tab stop at 2 inches (144 points) for the text in shape one on slide one in the active presentation.


```vb
Public Sub Ruler_Example() 
 
    Dim pptSlide As Slide 
    Set pptSlide = ActivePresentation.Slides(1) 
    pptSlide.Shapes(1).TextFrame2.Ruler.TabStops.Add ppTabStopLeft, 144 
 
End Sub
```


## See also


[TextFrame2 Object](PowerPoint.TextFrame2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]