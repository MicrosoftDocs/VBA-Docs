---
title: TextFrame2.Orientation property (PowerPoint)
keywords: vbapp10.chm678006
f1_keywords:
- vbapp10.chm678006
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.Orientation
ms.assetid: 713ce09e-575a-c1be-b60b-67884cb76673
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame2.Orientation property (PowerPoint)

 Returns or sets text orientation. Read/write.


## Syntax

_expression_.**Orientation**

 _expression_ An expression that returns a **[TextFrame2](PowerPoint.TextFrame2.md)** object.


## Return value

MsoTextOrientation


## Remarks

The value of the  **Orientation** property can be one of these **MsoTextOrientation** constants.


||
|:-----|
|**msoTextOrientationDownward**|
|**msoTextOrientationHorizontal**|
|**msoTextOrientationHorizontalRotatedFarEast**|
|**msoTextOrientationMixed**|
|**msoTextOrientationUpward**|
|**msoTextOrientationVertical**|
|**msoTextOrientationVerticalFarEast**|

> [!NOTE] 
> Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example shows how to orient the text horizontally in shape one on slide one in the active presentation.


```vb
Public Sub Orientation_Example()



    Dim pptSlide As Slide

    Set pptSlide = ActivePresentation.Slides(1)

    pptSlide.Shapes(1).TextFrame2.Orientation = msoTextOrientationHorizontal



End Sub
```


## See also


[TextFrame2 Object](PowerPoint.TextFrame2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]