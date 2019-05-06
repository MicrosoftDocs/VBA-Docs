---
title: CalendarSharing.IncludePrivateDetails property (Outlook)
keywords: vbaol11.chm2417
f1_keywords:
- vbaol11.chm2417
ms.prod: outlook
api_name:
- Outlook.CalendarSharing.IncludePrivateDetails
ms.assetid: a7c52e33-fe2a-b89a-9102-da2baf937e37
ms.date: 06/08/2017
localization_priority: Normal
---


# CalendarSharing.IncludePrivateDetails property (Outlook)

Returns or sets a  **Boolean** value that indicates whether private details for calendar items should be included in the iCalendar (.ics) file created by the **[ForwardAsICal](Outlook.CalendarSharing.ForwardAsICal.md)** or **[SaveAsICal](Outlook.CalendarSharing.SaveAsICal.md)** methods of the **[CalendarSharing](Outlook.CalendarSharing.md)** object. Read/write.


## Syntax

_expression_. `IncludePrivateDetails`

 _expression_ An expression that returns a [CalendarSharing](Outlook.CalendarSharing.md) object.


## Return value

 **True** if private details for calendar items should be included; otherwise, **False**.


## Remarks

This property must be set to  **False** if the **[CalendarDetail](Outlook.CalendarSharing.CalendarDetail.md)** property of the **CalendarSharing** object is set to **olFreeBusyOnly**.


## See also


[CalendarSharing Object](Outlook.CalendarSharing.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]