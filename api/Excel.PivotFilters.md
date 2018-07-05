---
title: PivotFilters Object (Excel)
keywords: vbaxl10.chm771072
f1_keywords:
- vbaxl10.chm771072
ms.prod: excel
api_name:
- Excel.PivotFilters
ms.assetid: fc647acb-bd6a-8544-6411-1f5e49807e53
ms.date: 06/08/2017
---


# PivotFilters Object (Excel)

The  **PivotFilters** object is a collection of **PivotFilter** objects.


## Remarks

The  **PivotFilters** collection contains properties and methods to add new filters, count the number of existing filters in the collection, and reference specific **PivotFilter** objects.


## Example

In the following example, a new PivotFilter is added to the PivotField at the currently active cell.


```vb
ActiveCell.PivotField.PivotFilters.Add FilterType := xlThisWeek
```


## Methods



|**Name**|
|:-----|
|[Add2](Excel.PivotFilters.Add.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.PivotFilters.Application.md)|
|[Count](Excel.PivotFilters.Count.md)|
|[Creator](Excel.PivotFilters.Creator.md)|
|[Item](Excel.PivotFilters.Item.md)|
|[Parent](Excel.PivotFilters.Parent.md)|

## See also


[Excel Object Model Reference](./overview/Excel/object-model.md)
