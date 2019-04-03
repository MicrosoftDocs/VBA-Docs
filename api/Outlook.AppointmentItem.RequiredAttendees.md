---
title: AppointmentItem.RequiredAttendees property (Outlook)
keywords: vbaol11.chm898
f1_keywords:
- vbaol11.chm898
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.RequiredAttendees
ms.assetid: 8ff112e9-2d8c-89de-0bdf-e8b9998f9269
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.RequiredAttendees property (Outlook)

Returns a semicolon-delimited  **String** of required attendee names for the meeting appointment. Read/write.


## Syntax

_expression_. `RequiredAttendees`

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Remarks

This property only contains the display names for the required attendees. The attendee list should be set by using the  **[Recipients](Outlook.Recipients.md)** collection.


## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]