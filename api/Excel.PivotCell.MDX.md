---
title: PivotCell.MDX property (Excel)
keywords: vbaxl10.chm692089
f1_keywords:
- vbaxl10.chm692089
ms.prod: excel
api_name:
- Excel.PivotCell.MDX
ms.assetid: 637dd366-5f83-e862-bab5-cf78db04a34e
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotCell.MDX property (Excel)

Returns a tuple that provides the full MDX coordinates of the specified value cell in a PivotTable with an OLAP data source. Read-only.


## Syntax

_expression_.**MDX**

_expression_ A variable that represents a **[PivotCell](Excel.PivotCell.md)** object.


## Return value

**String**


## Remarks

The dimensions returned in the tuple by the **MDX** property include row and column coordinates as well as report filter coordinates. 

For cells outside the values area of the PivotTable, and outside a PivotTable, accessing this property will generate a run-time error. 

For PivotTables with multi-item selection in a report filter field, accessing this property will also generate a run-time error.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]