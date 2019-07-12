---
title: ThreeDFormat.PresetLightingSoftness property (PowerPoint)
keywords: vbapp10.chm557013
f1_keywords:
- vbapp10.chm557013
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.PresetLightingSoftness
ms.assetid: 2dbe3666-2400-0142-01f8-995091f12311
ms.date: 06/08/2017
localization_priority: Normal
---


# ThreeDFormat.PresetLightingSoftness property (PowerPoint)

Returns or sets the intensity of the extrusion lighting. Read/write.


## Syntax

_expression_. `PresetLightingSoftness`

_expression_ A variable that represents a [ThreeDFormat](PowerPoint.ThreeDFormat.md) object.


## Return value

MsoPresetLightingSoftness


## Remarks

The value of the  **PresetLightingSoftness** property can be one of these **MsoPresetLightingSoftness** constants.


||
|:-----|
|**msoLightingBright**|
|**msoLightingDim**|
|**msoLightingNormal**|
|**msoPresetLightingSoftnessMixed**|

## Example

This example specifies that the extrusion for shape one on _myDocument_ be lit brightly from the left.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).ThreeD

    .Visible = True

    .PresetLightingSoftness = msoLightingBright

    .PresetLightingDirection = msoLightingLeft

End With
```


## See also


[ThreeDFormat Object](PowerPoint.ThreeDFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]