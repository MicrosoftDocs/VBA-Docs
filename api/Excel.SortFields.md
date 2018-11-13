---
title: SortFields object (Excel)
keywords: vbaxl10.chm844072
f1_keywords:
- vbaxl10.chm844072
ms.prod: excel
api_name:
- Excel.SortFields
ms.assetid: a9c83ea1-1cd9-1552-1f03-71bd92a2cc72
ms.date: 06/08/2017
---


# SortFields object (Excel)

The  **SortFields** collection is a collection of **SortField** objects. It allows developers to store a sort state on workbooks, lists, and autofilters.


## Remarks

The object contains properties to add, count, sort, and remove  **SortField** objects.


## Example


```vb
ActiveWorksheet.SortFields.Add Key:=Range("A1"), Order:=xlDescending 
ActiveWorksheet.SortFields.Add Key:=Range("B1"), Order:=xlDescending 
ActiveWorksheet.SortFields.Sort Header:=xlGuess 

```


## Methods



|**Name**|
|:-----|
|[Add](Excel.SortFields.Add.md)|
|[Clear](Excel.SortFields.Clear.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.SortFields.Application.md)|
|[Count](Excel.SortFields.Count.md)|
|[Creator](Excel.SortFields.Creator.md)|
|[Item](Excel.SortFields.Item.md)|
|[Parent](Excel.SortFields.Parent.md)|

## See also


[Excel Object Model Reference](./overview/Excel/object-model.md)
