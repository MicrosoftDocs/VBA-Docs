---
title: ValueChange.AllocationValue property (Excel)
keywords: vbaxl10.chm889078
f1_keywords:
- vbaxl10.chm889078
ms.prod: excel
api_name:
- Excel.ValueChange.AllocationValue
ms.assetid: 932cfa66-3664-5e23-85b7-769ac710669e
ms.date: 05/18/2019
localization_priority: Normal
---


# ValueChange.AllocationValue property (Excel)

Returns the value to allocate when performing what-if analysis. Read-only.


## Syntax

_expression_.**AllocationValue**

_expression_ A variable that represents a **[ValueChange](Excel.ValueChange.md)** object.


## Return value

**[XlAllocationValue](Excel.XlAllocationValue.md)**


## Remarks

The **AllocationValue** property corresponds to the **Value to Allocate** setting in the **What-If Analysis Settings** dialog box for a PivotTable report based on an OLAP data source as it was set at the time that this change was originally applied. 

If the specified **ValueChange** object was created by using the **[Add](Excel.PivotTableChangeList.Add.md)** method of the **PivotTableChangeList** collection, and the corresponding _AllocationValue_ parameter was not supplied, the default allocation value of the OLAP server is returned.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]