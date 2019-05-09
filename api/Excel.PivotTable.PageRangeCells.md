---
title: PivotTable.PageRangeCells property (Excel)
keywords: vbaxl10.chm235088
f1_keywords:
- vbaxl10.chm235088
ms.prod: excel
api_name:
- Excel.PivotTable.PageRangeCells
ms.assetid: 1c3b0694-539a-7d2d-17df-c0c0405d19e6
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.PageRangeCells property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents only the cells in the specified PivotTable report that contain the page fields and item drop-down lists.


## Syntax

_expression_.**PageRangeCells**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Example

This example selects only the cells in the PivotTable report that contain page fields and item drop-down lists.

```vb
Worksheets(1).PivotTables(1).PageRangeCells.Select
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]