---
title: Cell.Split method (PowerPoint)
keywords: vbapp10.chm628006
f1_keywords:
- vbapp10.chm628006
ms.prod: powerpoint
api_name:
- PowerPoint.Cell.Split
ms.assetid: edd81309-f0de-da70-67b2-4197059378fc
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.Split method (PowerPoint)

Splits a single table cell into multiple cells.


## Syntax

_expression_.**Split** (_NumRows_, _NumColumns_)

_expression_ A variable that represents a [Cell](PowerPoint.Cell.md) object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NumRows_|Required|**Long**|Number of rows that the cell is being split into.|
| _NumColumns_|Required|**Long**|Number of columns that the cell is being split into.|

## Example

This example splits the first cell in the referenced table into two cells, one directly above the other.


```vb
ActivePresentation.Slides(2).Shapes(5).Table.Cell(1, 1).Split 2, 1
```


## See also


[Cell Object](PowerPoint.Cell.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]