---
title: Shape.ShapeStyle property (Word)
keywords: vbawd10.chm161480854
f1_keywords:
- vbawd10.chm161480854
ms.prod: word
api_name:
- Word.Shape.ShapeStyle
ms.assetid: 7d6a6f4b-d55c-085e-1dab-c76ddbbb54ae
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.ShapeStyle property (Word)

Returns or sets the shape style for the specified shape. Read/write  **[MsoShapeStyleIndex](Office.MsoShapeStyleIndex.md)**.


## Syntax

_expression_.**ShapeStyle**

_expression_ A variable that represents a **[Shape](Word.Shape.md)** object.


## Example

The following code example changes the shape style for the first shape in the active document.


```vb
Dim myShape As Shape 
 
Set myShape = ActiveDocument.Shapes(1) 
 
myShape.ShapeStyle = msoLineStylePreset12
```


## See also


[Shape Object](Word.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]