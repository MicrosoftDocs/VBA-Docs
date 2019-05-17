---
title: ThreeDFormat.PresetLightingDirection property (Excel)
keywords: vbaxl10.chm119010
f1_keywords:
- vbaxl10.chm119010
ms.prod: excel
api_name:
- Excel.ThreeDFormat.PresetLightingDirection
ms.assetid: 5aea55a7-1718-a741-fc9b-f3e402469651
ms.date: 05/17/2019
localization_priority: Normal
---


# ThreeDFormat.PresetLightingDirection property (Excel)

Returns or sets the position of the light source relative to the extrusion. Read/write **[MsoPresetLightingDirection](office.msopresetlightingdirection.md)**.


## Syntax

_expression_.**PresetLightingDirection**

_expression_ A variable that represents a **[ThreeDFormat](Excel.ThreeDFormat.md)** object.


## Remarks

You won't see the lighting effects that you set if the extrusion has a wire frame surface.


## Example

This example specifies that the extrusion for shape one on _myDocument_ extend toward the top of the shape and that the lighting for the extrusion come from the left.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).ThreeD 
 .Visible = True 
 .SetExtrusionDirection msoExtrusionTop 
 .PresetLightingDirection = msoLightingLeft 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]