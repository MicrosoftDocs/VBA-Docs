---
title: ThreeDFormat object (Excel)
keywords: vbaxl10.chm120000
f1_keywords:
- vbaxl10.chm120000
ms.prod: excel
api_name:
- Excel.ThreeDFormat
ms.assetid: 9cb41236-6aba-4d6c-a54c-5e177657c8d1
ms.date: 04/02/2019
localization_priority: Normal
---


# ThreeDFormat object (Excel)

Represents a shape's three-dimensional formatting.


## Remarks

You cannot apply three-dimensional formatting to some kinds of shapes, such as beveled shapes or multiple-disjoint paths. Most of the properties and methods of the **ThreeDFormat** object for such a shape will fail.


## Example

Use the **[ThreeD](Excel.Shape.ThreeD.md)** property of the **Shape** object to return a **ThreeDFormat** object. The following example adds an oval to _myDocument_, and then specifies that the oval be extruded to a depth of 50 points and that the extrusion be purple.

```vb
Set myDocument = Worksheets(1) 
Set myShape = myDocument.Shapes.AddShape(msoShapeOval, _ 
 90, 90, 90, 40) 
With myShape.ThreeD 
 .Visible = True 
 .Depth = 50 
 .ExtrusionColor.RGB = RGB(255, 100, 255) 
 ' RGB value for purple 
End With
```
## Methods

- [IncrementRotationHorizontal](Excel.ThreeDFormat.IncrementRotationHorizontal.md)
- [IncrementRotationVertical](Excel.ThreeDFormat.IncrementRotationVertical.md)
- [IncrementRotationX](Excel.ThreeDFormat.IncrementRotationX.md)
- [IncrementRotationY](Excel.ThreeDFormat.IncrementRotationY.md)
- [IncrementRotationZ](Excel.ThreeDFormat.IncrementRotationZ.md)
- [ResetRotation](Excel.ThreeDFormat.ResetRotation.md)
- [SetExtrusionDirection](Excel.ThreeDFormat.SetExtrusionDirection.md)
- [SetPresetCamera](Excel.ThreeDFormat.SetPresetCamera.md)
- [SetThreeDFormat](Excel.ThreeDFormat.SetThreeDFormat.md)

## Properties

- [Application](Excel.ThreeDFormat.Application.md)
- [BevelBottomDepth](Excel.ThreeDFormat.BevelBottomDepth.md)
- [BevelBottomInset](Excel.ThreeDFormat.BevelBottomInset.md)
- [BevelBottomType](Excel.ThreeDFormat.BevelBottomType.md)
- [BevelTopDepth](Excel.ThreeDFormat.BevelTopDepth.md)
- [BevelTopInset](Excel.ThreeDFormat.BevelTopInset.md)
- [BevelTopType](Excel.ThreeDFormat.BevelTopType.md)
- [ContourColor](Excel.ThreeDFormat.ContourColor.md)
- [ContourWidth](Excel.ThreeDFormat.ContourWidth.md)
- [Creator](Excel.ThreeDFormat.Creator.md)
- [Depth](Excel.ThreeDFormat.Depth.md)
- [ExtrusionColor](Excel.ThreeDFormat.ExtrusionColor.md)
- [ExtrusionColorType](Excel.ThreeDFormat.ExtrusionColorType.md)
- [FieldOfView](Excel.ThreeDFormat.FieldOfView.md)
- [LightAngle](Excel.ThreeDFormat.LightAngle.md)
- [Parent](Excel.ThreeDFormat.Parent.md)
- [Perspective](Excel.ThreeDFormat.Perspective.md)
- [PresetCamera](Excel.ThreeDFormat.PresetCamera.md)
- [PresetExtrusionDirection](Excel.ThreeDFormat.PresetExtrusionDirection.md)
- [PresetLighting](Excel.ThreeDFormat.PresetLighting.md)
- [PresetLightingDirection](Excel.ThreeDFormat.PresetLightingDirection.md)
- [PresetLightingSoftness](Excel.ThreeDFormat.PresetLightingSoftness.md)
- [PresetMaterial](Excel.ThreeDFormat.PresetMaterial.md)
- [PresetThreeDFormat](Excel.ThreeDFormat.PresetThreeDFormat.md)
- [ProjectText](Excel.ThreeDFormat.ProjectText.md)
- [RotationX](Excel.ThreeDFormat.RotationX.md)
- [RotationY](Excel.ThreeDFormat.RotationY.md)
- [RotationZ](Excel.ThreeDFormat.RotationZ.md)
- [Visible](Excel.ThreeDFormat.Visible.md)
- [Z](Excel.ThreeDFormat.Z.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]