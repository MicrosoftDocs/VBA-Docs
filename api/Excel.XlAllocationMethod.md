---
title: XlAllocationMethod enumeration (Excel)
ms.prod: excel
api_name:
- Excel.xlAllocationMethod
ms.assetid: ade163bf-81d2-f633-323a-603b7c96e867
ms.date: 06/08/2017
localization_priority: Normal
---


# XlAllocationMethod enumeration (Excel)

Specifies the method to use to allocate values when performing what-if analysis on a PivotTable report based on an OLAP data source.

|Name|Value|Description|
|:-----|:-----|:-----|
| **xlEqualAllocation**|1|Use equal allocation.|
| **xlWeightedAllocation**|2|Use weighted allocation.|

## Remarks

If the **[AllocationMethod](Excel.PivotTable.AllocationMethod.md)** property is set to **xlWeightedAllocation**, you can optionally specify the weight expression to use by setting the **[AllocationWeightExpression](Excel.PivotTable.AllocationWeightExpression.md)** property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]