---
title: ThreeDFormat.SetExtrusionDirection method (Excel)
keywords: vbaxl10.chm119004
f1_keywords:
- vbaxl10.chm119004
ms.prod: excel
api_name:
- Excel.ThreeDFormat.SetExtrusionDirection
ms.assetid: 363c3150-fa6d-fcb3-d61d-00a36b528387
ms.date: 06/08/2017
---


# ThreeDFormat.SetExtrusionDirection method (Excel)

Sets the direction that the extrusion's sweep path takes away from the extruded shape.


## Syntax

 _expression_. `SetExtrusionDirection`( `_PresetExtrusionDirection_` )

 _expression_ A variable that represents a [ThreeDFormat](./Excel.ThreeDFormat.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PresetExtrusionDirection_|Required| **[MsoPresetExtrusionDirection](Office.MsoPresetExtrusionDirection.md)**|Specifies the extrusion direction.|

## Remarks



| **MsoPresetExtrusionDirection** can be one of these **MsoPresetExtrusionDirection** constants.|
| **msoExtrusionBottom**|
| **msoExtrusionBottomLeft**|
| **msoExtrusionBottomRight**|
| **msoExtrusionLeft**|
| **msoExtrusionNone**|
| **msoExtrusionRight**|
| **msoExtrusionTop**|
| **msoExtrusionTopLeft**|
| **msoExtrusionTopRight**|
| **msoPresetExtrusionDirectionMixed**|

This method sets the  **[PresetExtrusionDirection](Excel.ThreeDFormat.PresetExtrusionDirection.md)** property to the direction specified by the _PresetExtrusionDirection_ argument.


## Example

This example specifies that the extrusion for shape one on  `myDocument` extend toward the top of the shape and that the lighting for the extrusion come from the left.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).ThreeD 
    .Visible = True 
    .SetExtrusionDirection msoExtrusionTop 
    .PresetLightingDirection = msoLightingLeft 
End With
```


## See also


[ThreeDFormat Object](Excel.ThreeDFormat.md)

