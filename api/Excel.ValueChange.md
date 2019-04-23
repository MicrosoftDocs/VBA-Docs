---
title: ValueChange object (Excel)
keywords: vbaxl10.chm888072
f1_keywords:
- vbaxl10.chm888072
ms.prod: excel
api_name:
- Excel.ValueChange
ms.assetid: 27335d52-7003-2268-b5d0-c2cd21588579
ms.date: 04/03/2019
localization_priority: Normal
---


# ValueChange object (Excel)

Represents a value that has been changed in a PivotTable report that is based on an OLAP data source.


## Remarks

The **[PivotTableChangeList](Excel.PivotTableChangeList.md)** collection contains **ValueChange** objects that represent changes that a user has made to value cells in a PivotTable report.

The properties of the **ValueChange** object specify details about the change that was made, such as the value of the change, the tuple associated with the cell that was changed, the order in which the change was made relative to other changes, and whether the cell is visible in the PivotTable. 

The **ValueChange** object also provides the **PivotCell** property that returns a **[PivotCell](Excel.PivotCell.md)** object that represents the cell that was changed, and provides additional information about the changed cell.

## Methods

- [Delete](Excel.ValueChange.Delete.md)

## Properties

- [AllocationMethod](Excel.ValueChange.AllocationMethod.md)
- [AllocationValue](Excel.ValueChange.AllocationValue.md)
- [AllocationWeightExpression](Excel.ValueChange.AllocationWeightExpression.md)
- [Application](Excel.ValueChange.Application.md)
- [Creator](Excel.ValueChange.Creator.md)
- [Order](Excel.ValueChange.Order.md)
- [Parent](Excel.ValueChange.Parent.md)
- [PivotCell](Excel.ValueChange.PivotCell.md)
- [Tuple](Excel.ValueChange.Tuple.md)
- [Value](Excel.ValueChange.Value.md)
- [VisibleInPivotTable](Excel.ValueChange.VisibleInPivotTable.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]