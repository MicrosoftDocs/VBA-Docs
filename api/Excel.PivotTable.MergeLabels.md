---
title: PivotTable.MergeLabels property (Excel)
keywords: vbaxl10.chm235113
f1_keywords:
- vbaxl10.chm235113
ms.prod: excel
api_name:
- Excel.PivotTable.MergeLabels
ms.assetid: 2c658f34-1ec5-e1c8-59f7-b4401efc2646
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.MergeLabels property (Excel)

**True** if the specified PivotTable report's outer-row item, column item, subtotal, and grand total labels use merged cells. Read/write **Boolean**.


## Syntax

_expression_.**MergeLabels**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Example

This example causes the first PivotTable report on worksheet one to use merged-cell outer-row item, column item, subtotal, and grand total labels.

```vb
Worksheets(1).PivotTables(1).MergeLabels = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]