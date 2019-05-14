---
title: ShapeRange.ThreeD property (Excel)
keywords: vbaxl10.chm640116
f1_keywords:
- vbaxl10.chm640116
ms.prod: excel
api_name:
- Excel.ShapeRange.ThreeD
ms.assetid: 0b4ab4b8-841b-eea6-67a4-effe144d19fe
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.ThreeD property (Excel)

Returns a **[ThreeDFormat](Excel.ThreeDFormat.md)** object that contains 3D-effect formatting properties for the specified shape. Read-only.


## Syntax

_expression_.**ThreeD**

_expression_ A variable that represents a **[ShapeRange](Excel.shaperange.md)** object.


## Example

This example sets the depth, extrusion color, extrusion direction, and lighting direction for the 3D effects applied to shape one on _myDocument_.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).ThreeD 
 .Visible = True 
 .Depth = 50 
 .ExtrusionColor.RGB = RGB(255, 100, 255) 
 ' RGB value for purple 
 .SetExtrusionDirection msoExtrusionTop 
 .PresetLightingDirection = msoLightingLeft 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]