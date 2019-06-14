---
title: ShapeRange.ThreeD property (Publisher)
keywords: vbapb10.chm2293841
f1_keywords:
- vbapb10.chm2293841
ms.prod: publisher
api_name:
- Publisher.ShapeRange.ThreeD
ms.assetid: e5905f9d-dd84-b97e-ac5d-630f6c1208d7
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.ThreeD property (Publisher)

Returns a **[ThreeDFormat](Publisher.ThreeDFormat.md)** object.


## Syntax

_expression_.**ThreeD**

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Remarks

Use the **ThreeD** property to return a **ThreeDFormat** object whose properties are used to format the 3D appearance of the specified shape.


## Example

This example sets the depth, extrusion color, extrusion direction, and lighting direction for the 3D effects applied to shape one in the active publication.

```vb
Dim tdfTemp As ThreeDFormat 
 
Set tdfTemp = _ 
 ActiveDocument.Pages(1).Shapes(1).ThreeD 
 
With tdfTemp 
 .Visible = True 
 .Depth = 50 
 .ExtrusionColor.RGB = RGB(255, 100, 255) 
 .SetExtrusionDirection _ 
 PresetExtrusionDirection:=msoExtrusionTop 
 .PresetLightingDirection = msoLightingLeft 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]