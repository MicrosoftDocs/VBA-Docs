---
title: TimelineState object (Excel)
keywords: vbaxl10.chm949072
f1_keywords:
- vbaxl10.chm949072
ms.prod: excel
ms.assetid: bb92fe09-3cce-8e10-3795-2b9089c27801
ms.date: 04/02/2019
localization_priority: Normal
---


# TimelineState object (Excel)

The timeline-specific state of a **[SlicerCache](Excel.SlicerCache.md)** object.


## Remarks

Supported contiguous ranges can be set through the **SetFilterDateRange** method. When the timeline has such a contiguous filter state, the state can be retrieved from the two properties **StartDate** and **EndDate**. 

Any state that the filter may have, including non-contiguous states, can be retrieved through the three properties **FilterType**, **FilterValue1**, and **FilterValue2**.

## Methods

- [SetFilterDateRange](Excel.timelinestate.setfilterdaterange.md)

## Properties

- [Application](Excel.timelinestate.application.md)
- [Creator](Excel.timelinestate.creator.md)
- [EndDate](Excel.timelinestate.enddate.md)
- [FilterType](Excel.timelinestate.filtertype.md)
- [FilterValue1](Excel.timelinestate.filtervalue1.md)
- [FilterValue2](Excel.timelinestate.filtervalue2.md)
- [Parent](Excel.timelinestate.parent.md)
- [SingleRangeFilterState](Excel.timelinestate.singlerangefilterstate.md)
- [StartDate](Excel.timelinestate.startdate.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
