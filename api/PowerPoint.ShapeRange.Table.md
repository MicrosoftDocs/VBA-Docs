---
title: ShapeRange.Table property (PowerPoint)
keywords: vbapp10.chm548069
f1_keywords:
- vbapp10.chm548069
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Table
ms.assetid: 2ab10bd4-071a-8e84-cf46-1687e6661bb8
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.Table property (PowerPoint)

Returns a **[Table](PowerPoint.Table.md)** object that represents a table in a shape or in a shape range. Read-only.


## Syntax

_expression_. `Table`

_expression_ A variable that represents a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Return value

Table


## Example

This example sets the width of the first column in the table in shape five on the second slide to 80 points.


```vb
ActivePresentation.Slides(2).Shapes(5).Table.Columns(1).Width = 80
```


## See also


[ShapeRange Object](PowerPoint.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]