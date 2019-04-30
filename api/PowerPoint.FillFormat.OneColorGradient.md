---
title: FillFormat.OneColorGradient method (PowerPoint)
keywords: vbapp10.chm552003
f1_keywords:
- vbapp10.chm552003
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.OneColorGradient
ms.assetid: ce574185-2d13-993b-4a78-d681b6600621
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.OneColorGradient method (PowerPoint)

Sets the specified fill to a one-color gradient.


## Syntax

_expression_. `OneColorGradient`( `_Style_`, `_Variant_`, `_Degree_` )

_expression_ A variable that represents a **[FillFormat](powerpoint.fillformat.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Style_|Required|**MsoGradientStyle**|The gradient style.|
| _Variant_|Required|**Long**|The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the  **Gradient** tab in the **Shape Fill** tab. If Style is **msoGradientFromTitle** or **msoGradientFromCenter**, this argument can be either 1 or 2.|
| _Degree_|Required|**Single**|The gradient degree. Can be a value from 0.0 (dark) to 1.0 (light).|

## Remarks

The  _Style_ parameter value can be one of these **MsoGradientStyle** constants.


||
|:-----|
|**msoGradientDiagonalDown**|
|**msoGradientDiagonalUp**|
|**msoGradientFromCenter**|
|**msoGradientFromCorner**|
|**msoGradientFromTitle**|
|**msoGradientHorizontal**|
|**msoGradientMixed**|
|**msoGradientVertical**|

## Example

This example adds a rectangle with a one-color gradient fill to _myDocument_.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes _
        .AddShape(msoShapeRectangle, 90, 90, 90, 80).Fill
    .ForeColor.RGB = RGB(0, 128, 128)
    .OneColorGradient msoGradientHorizontal, 1, 1
End With
```


## See also


[FillFormat Object](PowerPoint.FillFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]