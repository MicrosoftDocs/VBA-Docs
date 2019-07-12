---
title: ThreeDFormat.PresetLightingDirection property (PowerPoint)
keywords: vbapp10.chm557012
f1_keywords:
- vbapp10.chm557012
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.PresetLightingDirection
ms.assetid: 85a5ae6c-5cdf-f4b5-ee9d-9ae220991037
ms.date: 06/08/2017
localization_priority: Normal
---


# ThreeDFormat.PresetLightingDirection property (PowerPoint)

Returns or sets the position of the light source relative to the extrusion. Read/write.


## Syntax

_expression_. `PresetLightingDirection`

_expression_ A variable that represents a [ThreeDFormat](PowerPoint.ThreeDFormat.md) object.


## Return value

MsoPresetLightingDirection


## Remarks

The value of the  **PresetLightingDirection** property can be one of these **MsoPresetLightingDirection** constants.


||
|:-----|
|**msoLightingBottom**|
|**msoLightingBottomLeft**|
|**msoLightingBottomRight**|
|**msoLightingLeft**|
|**msoLightingNone**|
|**msoLightingRight**|
|**msoLightingTop**|
|**msoLightingTopLeft**|
|**msoLightingTopRight**|
|**msoPresetLightingDirectionMixed**|

## Example

This example specifies that the extrusion for shape one on _myDocument_ extend toward the top of the shape and that the lighting for the extrusion come from the left.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).ThreeD

    .Visible = True

    .SetExtrusionDirection msoExtrusionTop

    .PresetLightingDirection = msoLightingLeft

End With
```


## See also


[ThreeDFormat Object](PowerPoint.ThreeDFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]