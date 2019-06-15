---
title: ThreeDFormat.SetExtrusionDirection method (Publisher)
keywords: vbapb10.chm3801108
f1_keywords:
- vbapb10.chm3801108
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.SetExtrusionDirection
ms.assetid: ac01d31d-7775-8e33-3b68-6e53f952fdda
ms.date: 06/15/2019
localization_priority: Normal
---


# ThreeDFormat.SetExtrusionDirection method (Publisher)

Sets the direction that the extrusion's sweep path takes away from the extruded shape.


## Syntax

_expression_.**SetExtrusionDirection** (_PresetExtrusionDirection_)

_expression_ A variable that represents a **[ThreeDFormat](Publisher.ThreeDFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_PresetExtrusionDirection_|Required| **[MsoPresetExtrusionDirection](office.msopresetextrusiondirection.md)**|Specifies the extrusion direction. Can be one of the **MsoPresetExtrusionDirection** constants declared in the Microsoft Office type library.|

## Remarks

This method sets the **[PresetExtrusionDirection](Publisher.ThreeDFormat.PresetExtrusionDirection.md)** property to the direction specified by the _PresetExtrusionDirection_ argument.


## Example

This example specifies that the extrusion for the first shape in the active publication extend toward the top of the shape and that the lighting for the extrusion come from the left.

```vb
With ActiveDocument.Pages(1).Shapes(1).ThreeD 
 .Visible = True 
 .SetExtrusionDirection _ 
 PresetExtrusionDirection:=msoExtrusionTop 
 .PresetLightingDirection = msoLightingLeft 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]