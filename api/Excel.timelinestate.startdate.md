---
title: TimelineState.StartDate property (Excel)
keywords: vbaxl10.chm950073
f1_keywords:
- vbaxl10.chm950073
ms.prod: excel
ms.assetid: 3de8df53-1a36-428e-50dd-c7f45aa73b25
ms.date: 05/18/2019
localization_priority: Normal
---


# TimelineState.StartDate property (Excel)

Returns the start of the filtering date range.  Read-only **Variant**.


## Syntax

_expression_.**StartDate**

_expression_ A variable that represents a **[TimelineState](Excel.TimelineState.md)** object.


## Remarks

This property returns an error for either of the following conditions:

- **[TimelineState.SingleRangeFilterState](Excel.timelinestate.singlerangefilterstate.md)** property == **False**
    
- **[SlicerCache.FilterCleared](Excel.slicercache.filtercleared.md)** property == **True**


## Property value

**VARIANT**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]