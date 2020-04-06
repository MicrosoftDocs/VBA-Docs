---
title: ShapeRange.Vertices property (Word)
keywords: vbawd10.chm162857086
f1_keywords:
- vbawd10.chm162857086
ms.prod: word
api_name:
- Word.ShapeRange.Vertices
ms.assetid: 1e27dbd8-2800-fe7f-4769-b6e9a4e802b5
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.Vertices property (Word)

Returns the coordinates of the specified freeform drawing's vertices (and control points for Bézier curves) as a series of coordinate pairs. You can use the array returned by this property as an argument for the  **AddCurve** or **AddPolyLine** method. Read-only **Variant**.


## Syntax

_expression_.**Vertices**

_expression_ Required. A variable that represents a **[ShapeRange](Word.shaperange.md)** object.


## Remarks

The following table shows how the  **Vertices** property associates values in the array _vertArray()_ with the coordinates of a triangle's vertices.



|**vertArray element**|**Contains**|
|:-----|:-----|
|
```vb
vertArray(1, 1)
```

|The horizontal distance from the first vertex to the left side of the document.|
|
```vb
vertArray(1, 2)
```

|The vertical distance from the first vertex to the top of the document.|
|
```vb
vertArray(2, 1)
```

|The horizontal distance from the second vertex to the left side of the document.|
|
```vb
vertArray(2, 2)
```

|The vertical distance from the second vertex to the top of the document.|
|
```vb
vertArray(3, 1)
```

|The horizontal distance from the third vertex to the left side of the document.|
|
```vb
vertArray(3, 2)
```

|The vertical distance from the third vertex to the top of the document.|

## See also


[ShapeRange Collection Object](Word.shaperange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]