---
title: ValueChange.Order property (Excel)
keywords: vbaxl10.chm889073
f1_keywords:
- vbaxl10.chm889073
ms.prod: excel
api_name:
- Excel.ValueChange.Order
ms.assetid: f64f8739-212b-6aca-3ddc-09c68c44978c
ms.date: 05/18/2019
localization_priority: Normal
---


# ValueChange.Order property (Excel)

Returns a value that indicates the order in which this change was performed relative to other changes in the **[PivotTableChangeList](Excel.PivotTableChangeList.md)** collection. Read-only.


## Syntax

_expression_.**Order**

_expression_ A variable that represents a **[ValueChange](Excel.ValueChange.md)** object.


## Return value

**Integer**


## Remarks

The value of the **Order** property is automatically assigned by Excel based on the order that the user applied the changes to value cells in the PivotTable report. If multiple changes were applied in one operation, Excel will arbitrarily assign the order within that set of changes.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]