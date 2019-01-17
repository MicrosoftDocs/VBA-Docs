---
title: PivotCache.RefreshDate property (Excel)
keywords: vbaxl10.chm227081
f1_keywords:
- vbaxl10.chm227081
ms.prod: excel
api_name:
- Excel.PivotCache.RefreshDate
ms.assetid: 0bbb3e62-584b-7daf-2ad0-643a6e886187
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotCache.RefreshDate property (Excel)

Returns the date on which the cache was last refreshed. Read-only  **Date**.


## Syntax

_expression_. `RefreshDate`

_expression_ A variable that represents a [PivotCache](Excel.PivotCache.md) object.


## Remarks

For  **PivotCache** objects, the cache must have at least one PivotTable report associated with it.

For OLAP data sources, this property is updated after each query.


## Example

This example displays the date on which the PivotTable report was last refreshed.


```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
dateString = Format(pvtTable.RefreshDate, "Long Date") 
MsgBox "The data was last refreshed on " & dateString
```


## See also


[PivotCache Object](Excel.PivotCache.md)

