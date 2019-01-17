---
title: TimelineState.StartDate property (Excel)
keywords: vbaxl10.chm950073
f1_keywords:
- vbaxl10.chm950073
ms.prod: excel
ms.assetid: 3de8df53-1a36-428e-50dd-c7f45aa73b25
ms.date: 06/08/2017
localization_priority: Normal
---


# TimelineState.StartDate property (Excel)

Returns the start of the filtering date range.  **Variant** Read-only


## Syntax

_expression_. `StartDate`

_expression_ A variable that represents a [TimelineState](Excel.timelinestate.md) object.


## Remarks

This property will return an error for either of the following conditions:


- [TimelineState.SingleRangeFilterState property (Excel)](Excel.timelinestate.singlerangefilterstate.md) == **False**
    
- [SlicerCache.FilterCleared property (Excel)](Excel.slicercache.filtercleared.md) == **True**
    

## Property value

 **VARIANT**


## See also



[TimelineState Object](Excel.timelinestate.md)

