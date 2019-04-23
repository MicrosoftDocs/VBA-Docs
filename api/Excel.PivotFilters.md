---
title: PivotFilters object (Excel)
keywords: vbaxl10.chm771072
f1_keywords:
- vbaxl10.chm771072
ms.prod: excel
api_name:
- Excel.PivotFilters
ms.assetid: fc647acb-bd6a-8544-6411-1f5e49807e53
ms.date: 03/30/2019
localization_priority: Normal
---


# PivotFilters object (Excel)

The **PivotFilters** object is a collection of **[PivotFilter](excel.pivotfilter.md)** objects.


## Remarks

The **PivotFilters** collection contains properties and methods to add new filters, count the number of existing filters in the collection, and reference specific **PivotFilter** objects.


## Example

In the following example, a new PivotFilter is added to the PivotField at the currently active cell.

```vb
ActiveCell.PivotField.PivotFilters.Add FilterType := xlThisWeek
```


## Methods

- [Add](Excel.PivotFilters.Add.md)

## Properties

- [Application](Excel.pivotFilters.Application.md)
- [Count](Excel.pivotFilters.Count.md)
- [Creator](Excel.pivotFilters.Creator.md)
- [Item](Excel.pivotFilters.Item.md)
- [Parent](Excel.pivotFilters.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]