---
title: TimelineState.EndDate property (Excel)
keywords: vbaxl10.chm950074
f1_keywords:
- vbaxl10.chm950074
ms.prod: excel
ms.assetid: 1d33ce70-32ed-a439-eb34-7305fd9557f2
ms.date: 05/18/2019
localization_priority: Normal
---


# TimelineState.EndDate property (Excel)

Returns the end of the filtering date range (equal to the **[StartDate](Excel.timelinestate.startdate.md)** property if range is a single day). Read-only **Variant**.


## Syntax

_expression_.**EndDate**

_expression_ A variable that represents a **[TimelineState](Excel.TimelineState.md)** object.


## Remarks

This property returns an error for either of the following conditions:

- **[TimelineState.SingleRangeFilterState](Excel.timelinestate.singlerangefilterstate.md)** property == **False**
    
- **[SlicerCache.FilterCleared](Excel.slicercache.filtercleared.md)** property == **True**
    

## Property value

**VARIANT**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]