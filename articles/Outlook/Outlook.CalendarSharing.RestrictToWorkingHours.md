---
title: CalendarSharing.RestrictToWorkingHours Property (Outlook)
keywords: vbaol11.chm2418
f1_keywords:
- vbaol11.chm2418
ms.prod: outlook
api_name:
- Outlook.CalendarSharing.RestrictToWorkingHours
ms.assetid: 2d655c66-fd3e-0b82-41b2-798d408f6531
ms.date: 06/08/2017
---


# CalendarSharing.RestrictToWorkingHours Property (Outlook)

Returns or sets a  **Boolean** value that indicates whether calendar items that do not occur within working hours should be included in the iCalendar (.ics) file created by the **[ForwardAsICal](Outlook.CalendarSharing.ForwardAsICal.md)** or **[SaveAsICal](Outlook.CalendarSharing.SaveAsICal.md)** methods of the **[CalendarSharing](Outlook.CalendarSharing.md)** object. Read/write.


## Syntax

 _expression_ . **RestrictToWorkingHours**

 _expression_ An expression that returns a **CalendarSharing** object.


### Return Value

 **True** if calendar items that do not occur within working hours should be included; otherwise, **False** .


## Remarks

This property must be set to  **False** if the **[CalendarDetail](Outlook.CalendarSharing.CalendarDetail.md)** property of the **CalendarSharing** object is set to **olFreeBusyOnly** or **olFullDetails** ..


## See also


#### Concepts


[CalendarSharing Object](Outlook.CalendarSharing.md)

