---
title: Cell.Borders property (PowerPoint)
keywords: vbapp10.chm628004
f1_keywords:
- vbapp10.chm628004
ms.prod: powerpoint
api_name:
- PowerPoint.Cell.Borders
ms.assetid: 1c9e2d38-237b-4c86-1135-af7533876501
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.Borders property (PowerPoint)

Returns a **[Borders](PowerPoint.Borders.md)** collection that represents the borders and diagonal lines for the specified **Cell** object or **CellRange** collection. Read-only.


## Syntax

_expression_.**Borders**

_expression_ A variable that represents a [Cell](PowerPoint.Cell.md) object.


## Return value

Borders


## Example

This example sets the thickness of the left border for the first cell in the second row of the selected table to three points.


```vb
ActiveWindow.Selection.ShapeRange.Table.Rows(2) _
    .Cells(1).Borders.Item(ppBorderLeft).Weight = 3
```


## See also


[Cell Object](PowerPoint.Cell.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]