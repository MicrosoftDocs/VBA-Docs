---
title: ThreeDFormat object (Publisher)
keywords: vbapb10.chm3866623
f1_keywords:
- vbapb10.chm3866623
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat
ms.assetid: 11d57330-c99e-5aa9-d47c-2c5d2846ed4d
ms.date: 06/04/2019
localization_priority: Normal
---


# ThreeDFormat object (Publisher)

Represents a shape's three-dimensional formatting.
 


## Remarks

You cannot apply three-dimensional formatting to some kinds of shapes, such as beveled shapes. Most of the properties and methods of the **ThreeDFormat** object for such a shape will fail.
 
Use the **[Shape.ThreeD](Publisher.Shape.ThreeD.md)** property to return a **ThreeDFormat** object. 
 

## Example

This example sets the depth, extrusion color, extrusion direction, and lighting direction for the 3D effects applied to shape one in the active publication.

```vb
Sub SetThreeDSettings() 
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
End Sub
```


## Methods

- [IncrementRotationX](Publisher.ThreeDFormat.IncrementRotationX.md)
- [IncrementRotationY](Publisher.ThreeDFormat.IncrementRotationY.md)
- [ResetRotation](Publisher.ThreeDFormat.ResetRotation.md)
- [SetExtrusionDirection](Publisher.ThreeDFormat.SetExtrusionDirection.md)
- [SetThreeDFormat](Publisher.ThreeDFormat.SetThreeDFormat.md)

## Properties

- [Application](Publisher.ThreeDFormat.Application.md)
- [BevelBottomDepth](Publisher.threedformat.bevelbottomdepth.md)
- [BevelBottomInset](Publisher.threedformat.bevelbottominset.md)
- [BevelBottomType](Publisher.threedformat.bevelbottomtype.md)
- [BevelTopDepth](Publisher.threedformat.beveltopdepth.md)
- [BevelTopInset](Publisher.threedformat.beveltopinset.md)
- [BevelTopType](Publisher.threedformat.beveltoptype.md)
- [ContourColor](Publisher.threedformat.contourcolor.md)
- [ContourWidth](Publisher.threedformat.contourwidth.md)
- [Depth](Publisher.ThreeDFormat.Depth.md)
- [ExtrusionColor](Publisher.ThreeDFormat.ExtrusionColor.md)
- [ExtrusionColorType](Publisher.ThreeDFormat.ExtrusionColorType.md)
- [FieldOfView](Publisher.threedformat.fieldofview.md)
- [Parent](Publisher.ThreeDFormat.Parent.md)
- [Perspective](Publisher.ThreeDFormat.Perspective.md)
- [PresetExtrusionDirection](Publisher.ThreeDFormat.PresetExtrusionDirection.md)
- [PresetLightingDirection](Publisher.ThreeDFormat.PresetLightingDirection.md)
- [PresetLightingSoftness](Publisher.ThreeDFormat.PresetLightingSoftness.md)
- [PresetMaterial](Publisher.ThreeDFormat.PresetMaterial.md)
- [PresetThreeDFormat](Publisher.ThreeDFormat.PresetThreeDFormat.md)
- [RotationX](Publisher.ThreeDFormat.RotationX.md)
- [RotationY](Publisher.ThreeDFormat.RotationY.md)
- [Visible](Publisher.ThreeDFormat.Visible.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]