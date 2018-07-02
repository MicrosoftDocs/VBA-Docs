---
title: ShapeRange.Vertices Property (Publisher)
keywords: vbapb10.chm2293845
f1_keywords:
- vbapb10.chm2293845
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Vertices
ms.assetid: 0beb2323-8db6-c8c2-2f34-4c1ffde7fddc
ms.date: 06/08/2017
---


# ShapeRange.Vertices Property (Publisher)

Returns the coordinates of the specified freeform drawing's vertices (and control points for Bézier curves) as a series of coordinate pairs. Read-only  **Variant**.


## Syntax

 _expression_. **Vertices**

 _expression_ A variable that represents a  **ShapeRange** object.


## Remarks

You can use the array returned by this property as an argument to the  [AddCurve](Publisher.Shapes.AddCurve.md)or  [AddPolyline](Publisher.Shapes.AddPolyline.md)methods.

The following table shows how the  **Vertices** property associates the values in the array `vertArray()` with the coordinates of a triangle's vertices.



|**vertArray element**|**Contains**|
|:-----|:-----|
| `vertArray(1, 1)`|The horizontal distance from the first vertex to the left side of the page.|
| `vertArray(1, 2)`|The vertical distance from the first vertex to the top of the page.|
| `vertArray(2, 1)`|The horizontal distance from the second vertex to the left side of the page.|
| `vertArray(2, 2)`|The vertical distance from the second vertex to the top of the page.|
| `vertArray(3, 1)`|The horizontal distance from the third vertex to the left side of the page.|
| `vertArray(3, 2)`|The vertical distance from the third vertex to the top of the page.|

## Example

This example assigns the vertex coordinates for shape one in the active publication to the array variable  `vertArray()` and displays the coordinates for the first vertex.


```vb
Dim vertArray As Variant 
Dim sngX1 As Single 
Dim sngY1 As Single 
 
With ActiveDocument.Pages(1).Shapes(1) 
 vertArray = .Vertices 
 sngX1 = vertArray(1, 1) 
 sngY1 = vertArray(1, 2) 
 MsgBox "First vertex coordinates: " &; sngX1 &; ", " &; sngY1 
End With
```

This example creates a curve that has the same geometric description as shape one in the active publication. Shape one must contain 3n+1 vertices for this example to work, where n is an integer greater than or equal to 1.




```vb
With ActiveDocument.Pages(1).Shapes 
 .AddCurve SafeArrayOfPoints:=.Item(1).Vertices 
End With 

```


