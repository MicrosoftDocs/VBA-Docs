---
title: Shape.Table property (PowerPoint)
keywords: vbapp10.chm547060
f1_keywords:
- vbapp10.chm547060
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Table
ms.assetid: cc57c50b-8c88-d863-31d2-a758eff5359f
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Table property (PowerPoint)

Returns a **[Table](PowerPoint.Table.md)** object that represents a table in a shape or in a shape range. Read-only.


## Syntax

_expression_. `Table`

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Return value

Table


## Example

This example sets the width of the first column in the table in shape five on the second slide to 80 points.


```vb
ActivePresentation.Slides(2).Shapes(5).Table.Columns(1).Width = 80
```


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]