---
title: CalendarSharing.CalendarDetail Property (Outlook)
keywords: vbaol11.chm2413
f1_keywords:
- vbaol11.chm2413
ms.prod: outlook
api_name:
- Outlook.CalendarSharing.CalendarDetail
ms.assetid: f3f0ba8d-23db-505f-58c4-6e3a33a468e7
ms.date: 06/08/2017
---


# CalendarSharing.CalendarDetail Property (Outlook)

Returns or sets an  **[OlCalendarDetail](Outlook.OlCalendarDetail.md)** value indicating the level of detail for calendar items included in the iCalendar (.ics) file created by the **[ForwardAsICal](Outlook.CalendarSharing.ForwardAsICal.md)** or **[SaveAsICal](Outlook.CalendarSharing.SaveAsICal.md)** methods of the **[CalendarSharing](Outlook.CalendarSharing.md)** object. Read/write.


## Syntax

 _expression_ . **CalendarDetail**

 _expression_ An expression that returns a **CalendarSharing** object.


### Return Value

A  **OlCalendarDetail** value that indicates the level of detail for calendar items.


## Remarks

The value of this property determines the allowable values for the following properties of the  **CalendarSharing** object:


-  **[IncludeAttachments](Outlook.CalendarSharing.IncludeAttachments.md)** must be set to **False** if **CalendarDetail** is set to **olFreeBusyOnly** or **olFreeBusyAndSubject** .
    
-  **[IncludePrivateDetails](Outlook.CalendarSharing.IncludePrivateDetails.md)** must be set to **False** if **CalendarDetail** is set to **olFreeBusyOnly** .
    
-  **[RestrictToWorkingHours](Outlook.CalendarSharing.RestrictToWorkingHours.md)** must be set to **False** if **CalendarDetail** is set to **olFreeBusyAndSubject** or **olFullDetails** .
    

## See also


#### Concepts


[CalendarSharing Object](Outlook.CalendarSharing.md)

