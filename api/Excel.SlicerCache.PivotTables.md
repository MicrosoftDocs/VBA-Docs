---
title: SlicerCache.PivotTables property (Excel)
keywords: vbaxl10.chm897078
f1_keywords:
- vbaxl10.chm897078
ms.prod: excel
api_name:
- Excel.SlicerCache.PivotTables
ms.assetid: 73fc8935-3c88-0a79-b0a1-05af99f14bc8
ms.date: 05/16/2019
localization_priority: Normal
---


# SlicerCache.PivotTables property (Excel)

Returns a **[SlicerPivotTables](Excel.SlicerPivotTables.md)** collection that contains information about the PivotTables that the slicer cache is currently filtering. Read-only.


## Syntax

_expression_.**PivotTables**

_expression_ A variable that represents a **[SlicerCache](Excel.SlicerCache.md)** object.


## Return value

**PivotTables**


## Remarks

The **SlicerPivotTables** collection returned by the **PivotTables** property will be empty if the slicer associated with the specified **SlicerCache** is not connected to any PivotTables.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]