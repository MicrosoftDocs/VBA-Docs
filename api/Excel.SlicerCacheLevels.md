---
title: SlicerCacheLevels object (Excel)
keywords: vbaxl10.chm898072
f1_keywords:
- vbaxl10.chm898072
ms.prod: excel
api_name:
- Excel.SlicerCacheLevels
ms.assetid: 6b1139a5-e81d-e11d-b4f5-f5d0fed24bf7
ms.date: 04/02/2019
localization_priority: Normal
---


# SlicerCacheLevels object (Excel)

Represents the collection of hierarchy levels for the OLAP data source that is filtered by a slicer.


## Remarks

When a slicer is used to filter an OLAP data source, its parent slicer cache can contain multiple hierarchy levels from the data source. 

Use the **SlicerCacheLevels** collection of the parent **[SlicerCache](Excel.SlicerCache.md)** object to access the **[SlicerCacheLevel](Excel.SlicerCacheLevel.md)** objects that represent these hierarchy levels. This collection is not accessible for non-OLAP data sources.


## Example

The following code example retrieves a **SlicerCacheLevel** object that represents the Country level of the Customer Geography hierarchy from the **SlicerCacheLevel** collection of the Country slicer.

```vb
ActiveWorkbook.SlicerCaches("Slicer_Customer_Geography"). _ 
 SlicerCacheLevels("[Customer].[Customer Geography].[Country]")
```

## Properties

- [Application](Excel.SlicerCacheLevels.Application.md)
- [Count](Excel.SlicerCacheLevels.Count.md)
- [Creator](Excel.SlicerCacheLevels.Creator.md)
- [Item](Excel.SlicerCacheLevels.Item.md)
- [Parent](Excel.SlicerCacheLevels.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]