---
title: Shapes.AddPolyline method (PowerPoint)
keywords: vbapp10.chm543011
f1_keywords:
- vbapp10.chm543011
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddPolyline
ms.assetid: e42c4f7a-de68-88bf-d250-28e642b56232
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddPolyline method (PowerPoint)

Creates an open polyline or a closed polygon drawing. Returns a  **[Shape](PowerPoint.Shape.md)** object that represents the new polyline or polygon.


## Syntax

_expression_. `AddPolyline`( `_SafeArrayOfPoints_` )

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SafeArrayOfPoints_|Required|**Variant**|An array of coordinate pairs that specifies the polyline drawing's vertices.|

## Return value

Shape


## Remarks

To form a closed polygon, assign the same coordinates to the first and last vertices in the polyline drawing.


## Example

This example adds a triangle to myDocument. Because the first and last points have the same coordinates, the polygon is closed and filled. The color of the triangle's interior will be the same as the default shape's fill color.


```vb
Dim triArray(1 To 4, 1 To 2) As Single

triArray(1, 1) = 25

triArray(1, 2) = 100

triArray(2, 1) = 100

triArray(2, 2) = 150

triArray(3, 1) = 150

triArray(3, 2) = 50

triArray(4, 1) = 25     ' Last point has same coordinates as first

triArray(4, 2) = 100

Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.AddPolyline SafeArrayOfPoints:=triArray
```


## See also


[Shapes Object](PowerPoint.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]