---
title: SlicerCache.SortItems property (Excel)
keywords: vbaxl10.chm897085
f1_keywords:
- vbaxl10.chm897085
ms.prod: excel
api_name:
- Excel.SlicerCache.SortItems
ms.assetid: da8fd267-5c4d-c333-fb21-bb3c4305747c
ms.date: 06/08/2017
localization_priority: Normal
---


# SlicerCache.SortItems property (Excel)

Returns or sets the sort order of the items in the slicer. Read/write  **[xlSlicerSort](Excel.XlSlicerSort.md)**.


## Syntax

_expression_. `SortItems`

_expression_ A variable that represents a '[SlicerCache](Excel.SlicerCache.md)' object.


## Return value

 **[xlSlicerSort](Excel.XlSlicerSort.md)**


## Remarks

The default setting of this property is  **xlSlicerSortAscending**.

The  **SortItems** property of the **SlicerCache** object only applies to slicers that are based on PivotTables that are connected to workbook ranges or lists (**SlicerCache**. **SourceType** = **xlDatabase**). Attempting to access the **SortItems** property for slicers that are connected to a OLAP data sources (**SlicerCache**. **[OLAP](Excel.SlicerCache.OLAP.md)** = **True**) generates a run-time error. For OLAP data sources, use the **[SortItems](Excel.SlicerCacheLevel.SortItems.md)** property of the **[SlicerCacheLevel](Excel.SlicerCacheLevel.md)** object instead.


## See also


[SlicerCache Object](Excel.SlicerCache.md)

