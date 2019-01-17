---
title: PivotTable.AllocationWeightExpression property (Excel)
keywords: vbaxl10.chm235190
f1_keywords:
- vbaxl10.chm235190
ms.prod: excel
api_name:
- Excel.PivotTable.AllocationWeightExpression
ms.assetid: 983f4819-5b3f-6f9d-667f-84feaf13bba5
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotTable.AllocationWeightExpression property (Excel)

Returns or sets the MDX weight expression to use when performing what-if analysis on a PivotTable report based on an OLAP data source. Read/write


## Syntax

_expression_. `AllocationWeightExpression`

_expression_ A variable that represents a '[PivotTable](Excel.PivotTable.md)' object.


## Remarks

The  **AllocationWeightExpression** property corresponds to the **Weight Expression** setting in the **What-If Analysis Settings** dialog box. Before the **AllocationWeightExpression** property can be set, you must set the **[AllocationMethod](Excel.PivotTable.AllocationMethod.md)** property to **xlWeightedAllocation**.


## See also


[PivotTable Object](Excel.PivotTable.md)

