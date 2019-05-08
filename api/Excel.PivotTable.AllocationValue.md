---
title: PivotTable.AllocationValue property (Excel)
keywords: vbaxl10.chm235188
f1_keywords:
- vbaxl10.chm235188
ms.prod: excel
api_name:
- Excel.PivotTable.AllocationValue
ms.assetid: c68351d8-2959-46db-1f43-ca1bc71e14fc
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.AllocationValue property (Excel)

Returns or sets the value to allocate when performing what-if analysis on a PivotTable report based on an OLAP data source. Read/write.


## Syntax

_expression_.**AllocationValue**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Return value

**[XlAllocationValue](Excel.XlAllocationValue.md)**


## Remarks

The **AllocationValue** property corresponds to the **Value to Allocate** setting in the **What-If Analysis Settings** dialog box. The default setting is **xlAllocateValue**, which corresponds to the setting **The value entered divided by the number of allocations**.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]