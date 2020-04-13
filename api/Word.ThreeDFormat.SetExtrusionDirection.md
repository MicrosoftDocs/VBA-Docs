---
title: ThreeDFormat.SetExtrusionDirection method (Word)
keywords: vbawd10.chm164626446
f1_keywords:
- vbawd10.chm164626446
ms.prod: word
api_name:
- Word.ThreeDFormat.SetExtrusionDirection
ms.assetid: 651b2b17-d87b-0007-3722-dc330f3e1f2e
ms.date: 06/08/2017
localization_priority: Normal
---


# ThreeDFormat.SetExtrusionDirection method (Word)

Sets the direction that the extrusion's sweep path takes away from the extruded shape.


## Syntax

_expression_. `SetExtrusionDirection`( `_PresetExtrusionDirection_` )

_expression_ Required. A variable that represents a '[ThreeDFormat](Word.ThreeDFormat.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PresetExtrusionDirection_|Required| **MsoPresetExtrusionDirection**|Sets the direction of the extrusion.|

## Remarks

This method sets the **PresetExtrusionDirection** property to the direction specified by the PresetExtrusionDirection argument.


## Example

This example specifies that the extrusion for the first shape on the active document extend toward the top of the shape and that the lighting for the extrusion come from the left.


```vb
With ActiveDocument.Shapes(1).ThreeD 
 .Visible = True 
 .SetExtrusionDirection msoExtrusionTop 
 .PresetLightingDirection = msoLightingLeft 
End With
```


## See also


[ThreeDFormat Object](Word.ThreeDFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]