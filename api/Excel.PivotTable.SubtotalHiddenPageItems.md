---
title: PivotTable.SubtotalHiddenPageItems property (Excel)
keywords: vbaxl10.chm235118
f1_keywords:
- vbaxl10.chm235118
ms.prod: excel
api_name:
- Excel.PivotTable.SubtotalHiddenPageItems
ms.assetid: bb3c7e54-1894-a1b6-e2d0-cf6097bd4875
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.SubtotalHiddenPageItems property (Excel)

**True** if hidden page field items in the PivotTable report are included in row and column subtotals, block totals, and grand totals. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**SubtotalHiddenPageItems**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

For OLAP data sources, the value is always **True**.


## Example

This example sets the first PivotTable report on worksheet one to exclude hidden page field items in subtotals.

```vb
Worksheets(1).PivotTables("Pivot1").SubtotalHiddenPageItems = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]