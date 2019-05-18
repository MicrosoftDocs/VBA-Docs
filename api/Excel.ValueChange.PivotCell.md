---
title: ValueChange.PivotCell property (Excel)
keywords: vbaxl10.chm889075
f1_keywords:
- vbaxl10.chm889075
ms.prod: excel
api_name:
- Excel.ValueChange.PivotCell
ms.assetid: 332859df-b643-cf9b-9c61-108f9324cee5
ms.date: 05/18/2019
localization_priority: Normal
---


# ValueChange.PivotCell property (Excel)

Returns a **[PivotCell](Excel.PivotCell.md)** object that represents the cell (tuple) that was changed. Read-only.


## Syntax

_expression_.**PivotCell**

_expression_ A variable that represents a **[ValueChange](Excel.ValueChange.md)** object.


## Return value

**PivotCell**


## Remarks

When the value of the **[VisibleInPivotTable](Excel.ValueChange.VisibleInPivotTable.md)** property of the specified **ValueChange** object is **True**, the **PivotCell** property returns a **PivotCell** object for the cell (tuple) that was changed. 

When the value of the **VisibleInPivotTable** property of the specified **ValueChange** object is **False**, the **PivotCell** property returns **NULL**.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]