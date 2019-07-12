---
title: ThreeDFormat.SetExtrusionDirection method (PowerPoint)
keywords: vbapp10.chm557006
f1_keywords:
- vbapp10.chm557006
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.SetExtrusionDirection
ms.assetid: 3ce76681-1a37-258b-594c-11d1d4f161c6
ms.date: 06/08/2017
localization_priority: Normal
---


# ThreeDFormat.SetExtrusionDirection method (PowerPoint)

Sets the direction that the extrusion's sweep path takes away from the extruded shape.


## Syntax

_expression_. `SetExtrusionDirection`( `_PresetExtrusionDirection_` )

_expression_ A variable that represents a [ThreeDFormat](PowerPoint.ThreeDFormat.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PresetExtrusionDirection_|Required|**MsoPresetExtrusionDirection**|Specifies the extrusion direction.|

## Remarks

This method sets the  **[PresetExtrusionDirection](PowerPoint.ThreeDFormat.PresetExtrusionDirection.md)** property to the direction specified by the PresetExtrusionDirection argument.

The PresetExtrusionDirection parameter value can be one of these  **MsoPresetExtrusionDirection** constants.


||
|:-----|
|**msoExtrusionBottom**|
|**msoExtrusionBottomLeft**|
|**msoExtrusionBottomRight**|
|**msoExtrusionLeft**|
|**msoExtrusionNone**|
|**msoExtrusionRight**|
|**msoExtrusionTop**|
|**msoExtrusionTopLeft**|
|**msoExtrusionTopRight**|
|**msoPresetExtrusionDirectionMixed**|

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