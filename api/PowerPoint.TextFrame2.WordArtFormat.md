---
title: TextFrame2.WordArtFormat property (PowerPoint)
keywords: vbapp10.chm678011
f1_keywords:
- vbapp10.chm678011
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.WordArtFormat
ms.assetid: 7ab4d90b-aae1-d98e-50d2-14b181d370ba
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame2.WordArtFormat property (PowerPoint)

Returns or sets the WordArt type for the specified text frame. Read/write.


## Syntax

_expression_. `WordArtFormat`

 _expression_ An expression that returns a **[TextFrame2](PowerPoint.TextFrame2.md)** object.


## Return value

MsoPresetTextEffect


## Remarks

The value of the  **WordArtFormat** property can be one of these **[MsoPresetTextEffect](Office.MsoPresetTextEffect.md)** constants.


## Example

This example shows how to set the word art format for shape one on slide one in the active presentation.


```vb
Public Sub WordArtFormat_Example() 
 
    Dim pptSlide As Slide 
    Set pptSlide = ActivePresentation.Slides(1) 
    pptSlide.Shapes(1).TextFrame2.WordArtFormat = msoTextEffect20 
     
End Sub
```


## See also


[TextFrame2 Object](PowerPoint.TextFrame2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]