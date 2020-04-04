---
title: Shapes.AddCurve method (PowerPoint)
keywords: vbapp10.chm543007
f1_keywords:
- vbapp10.chm543007
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddCurve
ms.assetid: 47f90182-a71b-a028-c43f-a85d59d2a56b
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddCurve method (PowerPoint)

Creates a Bézier curve. Returns a **[Shape](PowerPoint.Shape.md)** object that represents the new curve.


## Syntax

_expression_.**AddCurve** (_SafeArrayOfPoints_)

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SafeArrayOfPoints_|Required|**Variant**|An array of coordinate pairs that specifies the vertices and control points of the curve. The first point you specify is the starting vertex, and the next two points are control points for the first Bézier segment. Then, for each additional segment of the curve, you specify a vertex and two control points. The last point you specify is the ending vertex for the curve. Note that you must always specify 3n + 1 points, where n is the number of segments in the curve.|

## Return value

Shape


## Example

The following example adds a two-segment Bézier curve to myDocument.


```vb
Dim pts(1 To 7, 1 To 2) As Single

pts(1, 1) = 0

pts(1, 2) = 0

pts(2, 1) = 72

pts(2, 2) = 72

pts(3, 1) = 100

pts(3, 2) = 40

pts(4, 1) = 20

pts(4, 2) = 50

pts(5, 1) = 90

pts(5, 2) = 120

pts(6, 1) = 60

pts(6, 2) = 30

pts(7, 1) = 150

pts(7, 2) = 90

Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.AddCurve SafeArrayOfPoints:=pts
```


## See also


[Shapes Object](PowerPoint.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]