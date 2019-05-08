---
title: PivotTable.RowRange property (Excel)
keywords: vbaxl10.chm235095
f1_keywords:
- vbaxl10.chm235095
ms.prod: excel
api_name:
- Excel.PivotTable.RowRange
ms.assetid: 3b586599-9b2a-d0fc-c205-b8e3c6e7074f
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.RowRange property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents the range including the row area on the PivotTable report. Read-only.


## Syntax

_expression_.**RowRange**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Example

This example selects the row headers on the PivotTable report.

```vb
Worksheets("Sheet1").Activate 
Range("A3").Select 
ActiveCell.PivotTable.RowRange.Select
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]