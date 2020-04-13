---
title: CalendarSharing.IncludeWholeCalendar property (Outlook)
keywords: vbaol11.chm2420
f1_keywords:
- vbaol11.chm2420
ms.prod: outlook
api_name:
- Outlook.CalendarSharing.IncludeWholeCalendar
ms.assetid: 6cb75f0e-afb9-48fc-5b96-9f64a3b2ed6f
ms.date: 06/08/2017
localization_priority: Normal
---


# CalendarSharing.IncludeWholeCalendar property (Outlook)

Returns or sets a **Boolean** value that indicates whether all calendar items in the folder should be included in the iCalendar (.ics) file created by the **[ForwardAsICal](Outlook.CalendarSharing.ForwardAsICal.md)** or **[SaveAsICal](Outlook.CalendarSharing.SaveAsICal.md)** methods of the **[CalendarSharing](Outlook.CalendarSharing.md)** object. Read/write.


## Syntax

_expression_. `IncludeWholeCalendar`

 _expression_ An expression that returns a [CalendarSharing](Outlook.CalendarSharing.md) object.


## Return value

 **True** if all calendar items in the folder should be included; otherwise, **False**.


## Remarks

If this property is set to  **True**, the **[StartDate](Outlook.CalendarSharing.StartDate.md)** and **[EndDate](Outlook.CalendarSharing.EndDate.md)** properties of the **CalendarSharing** object are ignored and all calendar items in the folder are included.

If this property is set to  **False**, the **StartDate** and **EndDate** properties determine the range of calendar items to be included.


## See also


[CalendarSharing Object](Outlook.CalendarSharing.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]