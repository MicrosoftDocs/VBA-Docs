---
title: SlicerCache object (Excel)
keywords: vbaxl10.chm896072
f1_keywords:
- vbaxl10.chm896072
ms.prod: excel
api_name:
- Excel.SlicerCache
ms.assetid: 6e6533e3-0503-a1d3-9ecd-f7997233565f
ms.date: 04/02/2019
localization_priority: Normal
---


# SlicerCache object (Excel)

Represents the current filter state for a slicer, and information about which **[PivotCache](Excel.PivotCache.md)** or **[WorkbookConnection](Excel.WorkbookConnection.md)** object the slicer is connected to.


## Remarks

Use the **[SlicerCaches](Excel.Workbook.SlicerCaches.md)** property of the **Workbook** object to access the **[SlicerCaches](excel.slicercaches.md)** collection of **SlicerCache** objects in a workbook.

Each slicer has a base **SlicerCache** object that represents the items displayed in the slicer and the current user interface state of the tiles displayed with their corresponding item captions. Each slicer control that the user sees in Excel is represented by a **[Slicer](Excel.Slicer.md)** object that has a **SlicerCache** object associated with it.


## Example

The following code example creates a **SlicerCache** object based on the Customer Geography OLAP hierarchy from the connection to the AdventureWorks database, and then creates a slicer on the Country level of that hierarchy in Sheet2 of the workbook.

```vb
With ActiveWorkbook 
 .SlicerCaches.Add("AdventureWorks", _ 
 "[Customer].[Customer Geography]").Slicers.Add SlicerDestination:="Sheet2", _ 
 Level:="[Customer].[Customer Geography].[Country]", Caption:="Country" 
End With 

```


## Methods

- [ClearAllFilters](Excel.slicercache.clearallfilters.md)
- [ClearDateFilter](Excel.slicercache.cleardatefilter.md)
- [ClearManualFilter](Excel.slicercache.clearmanualfilter.md)
- [Delete](Excel.SlicerCache.Delete.md)

## Properties

- [Application](Excel.SlicerCache.Application.md)
- [Creator](Excel.SlicerCache.Creator.md)
- [CrossFilterType](Excel.SlicerCache.CrossFilterType.md)
- [FilterCleared](Excel.slicercache.filtercleared.md)
- [Index](Excel.SlicerCache.Index.md)
- [List](Excel.slicercache.list.md)
- [ListObject](Excel.slicercache.listobject.md)
- [Name](Excel.SlicerCache.Name.md)
- [OLAP](Excel.SlicerCache.OLAP.md)
- [Parent](Excel.SlicerCache.Parent.md)
- [PivotTables](Excel.SlicerCache.PivotTables.md)
- [RequireManualUpdate](Excel.slicercache.requiremanualupdate.md)
- [ShowAllItems](Excel.SlicerCache.ShowAllItems.md)
- [SlicerCacheLevels](Excel.SlicerCache.SlicerCacheLevels.md)
- [SlicerCacheType](Excel.slicercache.slicercachetype.md)
- [SlicerItems](Excel.SlicerCache.SlicerItems.md)
- [Slicers](Excel.SlicerCache.Slicers.md)
- [SortItems](Excel.SlicerCache.SortItems.md)
- [SortUsingCustomLists](Excel.SlicerCache.SortUsingCustomLists.md)
- [SourceName](Excel.SlicerCache.SourceName.md)
- [SourceType](Excel.SlicerCache.SourceType.md)
- [TimelineState](Excel.slicercache.timelinestate.md)
- [VisibleSlicerItems](Excel.SlicerCache.VisibleSlicerItems.md)
- [VisibleSlicerItemsList](Excel.SlicerCache.VisibleSlicerItemsList.md)
- [WorkbookConnection](Excel.SlicerCache.WorkbookConnection.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
