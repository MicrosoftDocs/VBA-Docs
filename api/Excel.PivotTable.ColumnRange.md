---
title: PivotTable.ColumnRange property (Excel)
keywords: vbaxl10.chm235076
f1_keywords:
- vbaxl10.chm235076
ms.prod: excel
api_name:
- Excel.PivotTable.ColumnRange
ms.assetid: 7f54b908-b0cb-80c8-e16f-25c7ff536e43
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.ColumnRange property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents the range that contains the column area in the PivotTable report. Read-only.


## Syntax

_expression_.**ColumnRange**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Example

This example selects the column headers for the PivotTable report.

```vb
Worksheets("Sheet1").Activate 
Range("A3").Select 
ActiveCell.PivotTable.ColumnRange.Select
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]