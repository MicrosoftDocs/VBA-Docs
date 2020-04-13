---
title: CalendarSharing.IncludeAttachments property (Outlook)
keywords: vbaol11.chm2416
f1_keywords:
- vbaol11.chm2416
ms.prod: outlook
api_name:
- Outlook.CalendarSharing.IncludeAttachments
ms.assetid: 504bba9e-009f-986f-070e-ff73ce82ea03
ms.date: 06/08/2017
localization_priority: Normal
---


# CalendarSharing.IncludeAttachments property (Outlook)

Returns or sets a **Boolean** value that indicates whether attachments for calendar items should be included in the iCalendar (.ics) file created by the **[ForwardAsICal](Outlook.CalendarSharing.ForwardAsICal.md)** or **[SaveAsICal](Outlook.CalendarSharing.SaveAsICal.md)** methods of the **[CalendarSharing](Outlook.CalendarSharing.md)** object. Read/write.


## Syntax

_expression_. `IncludeAttachments`

 _expression_ An expression that returns a [CalendarSharing](Outlook.CalendarSharing.md) object.


## Return value

 **True** if attachments for calendar items should be included; otherwise, **False**.


## Remarks

This property must be set to  **False** if the **[CalendarDetail](Outlook.CalendarSharing.CalendarDetail.md)** property of the **CalendarSharing** object is set to **olFreeBusyOnly** or **olFreeBusyAndSubject**.


## See also


[CalendarSharing Object](Outlook.CalendarSharing.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]