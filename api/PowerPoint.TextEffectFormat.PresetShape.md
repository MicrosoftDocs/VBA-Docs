---
title: TextEffectFormat.PresetShape property (PowerPoint)
keywords: vbapp10.chm556010
f1_keywords:
- vbapp10.chm556010
ms.prod: powerpoint
api_name:
- PowerPoint.TextEffectFormat.PresetShape
ms.assetid: e4e43c4c-09fa-4f1d-a0de-26e0c7a872a0
ms.date: 06/08/2017
localization_priority: Normal
---


# TextEffectFormat.PresetShape property (PowerPoint)

Returns or sets the shape of the specified WordArt. Read/write.


## Syntax

_expression_. `PresetShape`

_expression_ A variable that represents a [TextEffectFormat](PowerPoint.TextEffectFormat.md) object.


## Return value

MsoPresetTextEffectShape


## Remarks

Setting the  **[PresetTextEffect](PowerPoint.TextEffectFormat.PresetTextEffect.md)** property automatically sets the **PresetShape** property.

The value of the  **PresetShape** property can be one of these **MsoPresetTextEffectShape** constants.


||
|:-----|
|**msoTextEffectShapeArchDownCurve**|
|**msoTextEffectShapeArchDownPour**|
|**msoTextEffectShapeArchUpCurve**|
|**msoTextEffectShapeArchUpPour**|
|**msoTextEffectShapeButtonCurve**|
|**msoTextEffectShapeButtonPour**|
|**msoTextEffectShapeCanDown**|
|**msoTextEffectShapeCanUp**|
|**msoTextEffectShapeCascadeDown**|
|**msoTextEffectShapeCascadeUp**|
|**msoTextEffectShapeChevronDown**|
|**msoTextEffectShapeChevronUp**|
|**msoTextEffectShapeCircleCurve**|
|**msoTextEffectShapeCirclePour**|
|**msoTextEffectShapeCurveDown**|
|**msoTextEffectShapeCurveUp**|
|**msoTextEffectShapeDeflate**|
|**msoTextEffectShapeDeflateBottom**|
|**msoTextEffectShapeDeflateInflate**|
|**msoTextEffectShapeDeflateInflateDeflate**|
|**msoTextEffectShapeDeflateTop**|
|**msoTextEffectShapeDoubleWave2**|
|**msoTextEffectShapeFadeDown**|
|**msoTextEffectShapeFadeLeft**|
|**msoTextEffectShapeFadeRight**|
|**msoTextEffectShapeFadeUp**|
|**msoTextEffectShapeInflate**|
|**msoTextEffectShapeInflateBottom**|
|**msoTextEffectShapeInflateTop**|
|**msoTextEffectShapeMixed**|
|**msoTextEffectShapePlainText**|
|**msoTextEffectShapeRingInside**|
|**msoTextEffectShapeRingOutside**|
|**msoTextEffectShapeSlantDown**|
|**msoTextEffectShapeSlantUp**|
|**msoTextEffectShapeStop**|
|**msoTextEffectShapeTriangleDown**|
|**msoTextEffectShapeTriangleUp**|
|**msoTextEffectShapeWave1**|
|**msoTextEffectShapeWave2**|
|**msoTextEffectShapeDoubleWave1**|

## Example

This example sets the shape of all WordArt on _myDocument_ to a chevron whose center points down.


```vb
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Shapes

    If s.Type = msoTextEffect Then

        s.TextEffect.PresetShape = msoTextEffectShapeChevronDown

    End If

Next
```


## See also


[TextEffectFormat Object](PowerPoint.TextEffectFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]