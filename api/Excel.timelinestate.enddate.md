---
title: TimelineState.EndDate Property (Excel)
keywords: vbaxl10.chm950074
f1_keywords:
- vbaxl10.chm950074
ms.prod: excel
ms.assetid: 1d33ce70-32ed-a439-eb34-7305fd9557f2
ms.date: 06/08/2017
---


# TimelineState.EndDate Property (Excel)

Returns the end of the filtering date range (equals to [TimelineState.StartDate Property (Excel)](Excel.timelinestate.startdate.md) if range is a single day). **Variant** Read-only


## Syntax

 _expression_. `EndDate`

 _expression_ A variable that represents a[TimelineState](Excel.timelinestate.md) object.


## Remarks

This property will return an error for either of the following conditions:


- [TimelineState.SingleRangeFilterState Property (Excel)](Excel.timelinestate.singlerangefilterstate.md) == **False**
    
- [SlicerCache.FilterCleared Property (Excel)](Excel.slicercache.filtercleared.md) == **True**
    

## Property value

 **VARIANT**


## See also



[TimelineState Object](Excel.timelinestate.md)

