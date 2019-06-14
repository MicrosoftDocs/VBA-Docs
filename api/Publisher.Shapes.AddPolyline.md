---
title: Shapes.AddPolyline method (Publisher)
keywords: vbapb10.chm2162711
f1_keywords:
- vbapb10.chm2162711
ms.prod: publisher
api_name:
- Publisher.Shapes.AddPolyline
ms.assetid: d49fb2bc-4df5-fff8-c741-2c0d35413fc5
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.AddPolyline method (Publisher)

Adds a new **[Shape](Publisher.Shape.md)** object representing an open polyline or a closed polygon to the specified **Shapes** collection.


## Syntax

_expression_.**AddPolyline** (_SafeArrayOfPoints_)

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_SafeArrayOfPoints_|Required| **Variant**|An array of coordinate pairs that specifies the polyline's or polygon's vertices.|

## Return value

Shape


## Remarks

For the array elements in _SafeArrayOfPoints_, numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

To form a closed polygon, assign the same coordinates to the first and last vertices in the polyline drawing.


## Example

The following example adds a triangle to the first page of the active publication. Because the first and last points have the same coordinates, the polygon is closed.

```vb
Dim shpPolyline As Shape 
Dim arrPoints(1 To 4, 1 To 2) As Single 
 
arrPoints(1, 1) = 25 
arrPoints(1, 2) = 100 
arrPoints(2, 1) = 100 
arrPoints(2, 2) = 150 
arrPoints(3, 1) = 150 
arrPoints(3, 2) = 50 
arrPoints(4, 1) = 25 
arrPoints(4, 2) = 100 
 
Set shpPolyline = ActiveDocument.Pages(1).Shapes.AddPolyline _ 
 (SafeArrayOfPoints:=arrPoints)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]