---
title: PivotTable.AllocationMethod property (Excel)
keywords: vbaxl10.chm235189
f1_keywords:
- vbaxl10.chm235189
ms.prod: excel
api_name:
- Excel.PivotTable.AllocationMethod
ms.assetid: 726393d4-4aba-556a-9278-976e7b9a1088
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.AllocationMethod property (Excel)

Returns or sets the method to use to allocate values when performing what-if analysis on a PivotTable report based on an OLAP data source. Read/write.


## Syntax

_expression_.**AllocationMethod**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Return value

**[XlAllocationMethod](Excel.XlAllocationMethod.md)**


## Remarks

The **AllocationMethod** property corresponds to the **Allocation Method** setting in the **What-If Analysis Settings** dialog box. The default setting is **xlEqualAllocation**, which corresponds to the **Equal Allocation** setting. 

If the **AllocationMethod** property is set to **xlWeightedAllocation**, which corresponds to the **Weighted Allocation** setting, you can optionally specify the weight expression to use by setting the **[AllocationWeightExpression](Excel.PivotTable.AllocationWeightExpression.md)** property. 

If you do not specify a weight expression, a weight expression equivalent to `<leaf cell value> / <existing value>` is used.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]