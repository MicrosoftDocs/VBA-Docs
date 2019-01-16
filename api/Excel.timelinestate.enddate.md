---
title: TimelineState.EndDate property (Excel)
keywords: vbaxl10.chm950074
f1_keywords:
- vbaxl10.chm950074
ms.prod: excel
ms.assetid: 1d33ce70-32ed-a439-eb34-7305fd9557f2
ms.date: 06/08/2017
localization_priority: Normal
---


# TimelineState.EndDate property (Excel)

Returns the end of the filtering date range (equals to [TimelineState.StartDate property (Excel)](Excel.timelinestate.startdate.md) if range is a single day). **Variant** Read-only


## Syntax

_expression_. `EndDate`

_expression_ A variable that represents a [TimelineState](Excel.timelinestate.md) object.


## Remarks

This property will return an error for either of the following conditions:


- [TimelineState.SingleRangeFilterState property (Excel)](Excel.timelinestate.singlerangefilterstate.md) == **False**
    
- [SlicerCache.FilterCleared property (Excel)](Excel.slicercache.filtercleared.md) == **True**
    

## Property value

 **VARIANT**


## See also



[TimelineState Object](Excel.timelinestate.md)

