---
title: ThreeDFormat object (Excel)
keywords: vbaxl10.chm120000
f1_keywords:
- vbaxl10.chm120000
ms.prod: excel
api_name:
- Excel.ThreeDFormat
ms.assetid: 9cb41236-6aba-4d6c-a54c-5e177657c8d1
ms.date: 06/08/2017
localization_priority: Normal
---


# ThreeDFormat object (Excel)

Represents a shape's three-dimensional formatting.


## Remarks

You cannot apply three-dimensional formatting to some kinds of shapes, such as beveled shapes or multiple-disjoint paths. Most of the properties and methods of the  **ThreeDFormat** object for such a shape will fail.


## Example

Use the  **[ThreeD](Excel.Shape.ThreeD.md)** property to return a **ThreeDFormat** object. The following example adds an oval to _myDocument_ and then specifies that the oval be extruded to a depth of 50 points and that the extrusion be purple.


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


## See also



[Excel Object Model Reference](./overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]