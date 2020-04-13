---
title: CalendarSharing.EndDate property (Outlook)
keywords: vbaol11.chm2414
f1_keywords:
- vbaol11.chm2414
ms.prod: outlook
api_name:
- Outlook.CalendarSharing.EndDate
ms.assetid: 89358c71-7805-7acc-5afb-2ba7b592f9f2
ms.date: 06/08/2017
localization_priority: Normal
---


# CalendarSharing.EndDate property (Outlook)

Returns or sets a **Date** value that represents the inclusive end date of the range of calendar items to be shared by the **[CalendarSharing](Outlook.CalendarSharing.md)** object. Read/write.


## Syntax

_expression_. `EndDate`

 _expression_ An expression that returns a [CalendarSharing](Outlook.CalendarSharing.md) object.


## Return value

A **Date** value representing the inclusive end date of the range of calendar items to be shared.


## Remarks

This property is ignored if the  **[IncludeWholeCalendar](Outlook.CalendarSharing.IncludeWholeCalendar.md)** property of the **CalendarSharing** object is set to **True**.


## See also


[CalendarSharing Object](Outlook.CalendarSharing.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]