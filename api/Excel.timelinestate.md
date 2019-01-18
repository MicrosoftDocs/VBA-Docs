---
title: TimelineState object (Excel)
keywords: vbaxl10.chm949072
f1_keywords:
- vbaxl10.chm949072
ms.prod: excel
ms.assetid: bb92fe09-3cce-8e10-3795-2b9089c27801
ms.date: 06/08/2017
localization_priority: Normal
---


# TimelineState object (Excel)

The timeline-specific state of a [SlicerCache object (Excel)](Excel.SlicerCache.md) object.


## Remarks

Supported contiguous ranges can be set through the [TimelineState.SetFilterDateRange method (Excel)](Excel.timelinestate.setfilterdaterange.md) method. When the Timeline has such a contiguous filter state, the state can be retrieved from the two properties[TimelineState.StartDate property (Excel)](Excel.timelinestate.startdate.md) and [TimelineState.EndDate property (Excel)](Excel.timelinestate.enddate.md). Any state that the filter may have, including non-contiguous states, can be retrieved through the three properties: [TimelineState.FilterType property (Excel)](Excel.timelinestate.filtertype.md), [TimelineState.FilterValue1 property (Excel)](Excel.timelinestate.filtervalue1.md), and [TimelineState.FilterValue2 property (Excel)](Excel.timelinestate.filtervalue2.md).


## See also

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]