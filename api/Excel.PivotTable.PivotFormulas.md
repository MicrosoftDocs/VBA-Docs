---
title: PivotTable.PivotFormulas property (Excel)
keywords: vbaxl10.chm235116
f1_keywords:
- vbaxl10.chm235116
ms.prod: excel
api_name:
- Excel.PivotTable.PivotFormulas
ms.assetid: fceade1d-7aa1-85c1-ca74-89460ffa6dff
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.PivotFormulas property (Excel)

Returns a **[PivotFormulas](Excel.PivotFormulas.md)** object that represents the collection of formulas for the specified PivotTable report. Read-only.


## Syntax

_expression_.**PivotFormulas**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

For OLAP data sources, this property returns an empty collection.


## Example

```vb
For Each pf in ActiveSheet.PivotTables(1).PivotFormulas 
 r = r + 1 
 Cells(r, 1).Value = pf.Formula 
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]