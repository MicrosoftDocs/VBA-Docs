---
title: ThreeDFormat.SetThreeDFormat method (Publisher)
keywords: vbapb10.chm3801107
f1_keywords:
- vbapb10.chm3801107
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.SetThreeDFormat
ms.assetid: d73dbada-1a33-4b5b-9733-e228a0cc5f8c
ms.date: 06/15/2019
localization_priority: Normal
---


# ThreeDFormat.SetThreeDFormat method (Publisher)

Sets the preset extrusion format. Each preset extrusion format contains a set of preset values for the 3D properties of the extrusion.


## Syntax

_expression_.**SetThreeDFormat** (_PresetThreeDFormat_)

_expression_ A variable that represents a **[ThreeDFormat](Publisher.ThreeDFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_PresetThreeDFormat_|Required| **[MsoPresetThreeDFormat](Office.MsoPresetThreeDFormat.md)**|Specifies a preset extrusion format that corresponds to one of the options (numbered from left to right, from top to bottom) displayed when you choose the **3D** button on the **Drawing** toolbar. Can be one of the **MsoPresetThreeDFormat** constants declared in the Microsoft Office type library.|

## Remarks

This method sets the **[PresetThreeDFormat](Publisher.ThreeDFormat.PresetThreeDFormat.md)** property to the format specified by the _PresetThreeDFormat_ argument.


## Example

This example adds an oval to the active publication and sets its extrusion format to one of the preset 3D formats.

```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeOval, _ 
 Left:=30, Top:=30, Width:=50, Height:=25).ThreeD 
 .Visible = True 
 .SetThreeDFormat PresetThreeDFormat:=msoThreeD12 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]