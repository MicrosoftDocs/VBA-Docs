---
title: ThreeDFormat.PresetExtrusionDirection property (Excel)
keywords: vbaxl10.chm119009
f1_keywords:
- vbaxl10.chm119009
ms.prod: excel
api_name:
- Excel.ThreeDFormat.PresetExtrusionDirection
ms.assetid: 61f75976-03d4-b449-31e3-e0c7839cce92
ms.date: 05/17/2019
localization_priority: Normal
---


# ThreeDFormat.PresetExtrusionDirection property (Excel)

Returns the direction that the extrusion's sweep path takes away from the extruded shape (the front face of the extrusion). Read-only **[MsoPresetExtrusionDirection](office.msopresetextrusiondirection.md)**.

## Syntax

_expression_.**PresetExtrusionDirection**

_expression_ A variable that represents a **[ThreeDFormat](Excel.ThreeDFormat.md)** object.


## Remarks

This property is read-only. To set the value of this property, use the **[SetExtrusionDirection](Excel.ThreeDFormat.SetExtrusionDirection.md)** method.


## Example

This example changes each extrusion on _myDocument_ that extends toward the upper-left corner of the extrusion's front face to an extrusion that extends toward the lower-right corner of the front face.

```vb
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
 With s.ThreeD 
 If .PresetExtrusionDirection = msoExtrusionTopLeft Then 
 .SetExtrusionDirection msoExtrusionBottomRight 
 End If 
 End With 
Next
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]