---
title: ValueChange.AllocationMethod property (Excel)
keywords: vbaxl10.chm889079
f1_keywords:
- vbaxl10.chm889079
ms.prod: excel
api_name:
- Excel.ValueChange.AllocationMethod
ms.assetid: 124ff77d-56f0-7877-a4ed-9c62e1d217d1
ms.date: 05/18/2019
localization_priority: Normal
---


# ValueChange.AllocationMethod property (Excel)

Returns the method to use to allocate this value when performing what-if analysis. Read-only.


## Syntax

_expression_.**AllocationMethod**

_expression_ A variable that represents a **[ValueChange](Excel.ValueChange.md)** object.


## Return value

**[XlAllocationMethod](Excel.XlAllocationMethod.md)**


## Remarks

The **AllocationMethod** property corresponds to the **Allocation Method** setting in the **What-If Analysis Settings** dialog box for a PivotTable report based on an OLAP data source as it was set at the time that this change was originally applied. 

If the specified **ValueChange** object was created by using the **[Add](Excel.PivotTableChangeList.Add.md)** method of the **PivotTableChangeList** collection, and the corresponding _AllocationMethod_ parameter was not supplied, the default allocation method of the OLAP server is returned.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]