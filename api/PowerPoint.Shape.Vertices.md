---
title: Shape.Vertices Property (PowerPoint)
keywords: vbapp10.chm547040
f1_keywords:
- vbapp10.chm547040
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Vertices
ms.assetid: b9ce441c-b305-4224-3fe8-3f947199a4d4
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Vertices Property (PowerPoint)

Returns the coordinates of the specified freeform drawing's vertices (and control points for B?zier curves) as a series of coordinate pairs. Read-only.


## Syntax

 _expression_. `Vertices`

 _expression_ A variable that represents a [Shape](./PowerPoint.Shape.md) object.


## Return value

Variant


## Remarks

You can use the array returned by this property as an argument to the  **[AddCurve](PowerPoint.Shapes.AddCurve.md)** method or **[AddPolyline](PowerPoint.Shapes.AddPolyline.md)** method.

The following table shows how the  **Vertices** property associates the values in the array `vertArray()` with the coordinates of a triangle's vertices.



|**VertArray element**|**Contains**|
|:-----|:-----|
| `VertArray(1, 1)`|The horizontal distance from the first vertex to the left side of the slide|
| `VertArray(1, 2)`|The vertical distance from the first vertex to the top of the slide|
| `VertArray(2, 1)`|The horizontal distance from the second vertex to the left side of the slide|
| `VertArray(2, 2)`|The vertical distance from the second vertex to the top of the slide|
| `VertArray(3, 1)`|The horizontal distance from the third vertex to the left side of the slide|
| `VertArray(3, 2)`|The vertical distance from the third vertex to the top of the slide|

## Example

This example assigns the vertex coordinates for shape one on  `myDocument` to the array variable `vertArray()` and displays the coordinates for the first vertex.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1)

    vertArray = .Vertices

    x1 = vertArray(1, 1)

    y1 = vertArray(1, 2)

    MsgBox "First vertex coordinates: " & x1 & ", " & y1

End With
```

This example creates a curve that has the same geometric description as shape one on  `myDocument`. Shape one must contain 3 _n_ +1 vertices for this example to succeed.




```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes

    .AddCurve .Item(1).Vertices

End With
```


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]