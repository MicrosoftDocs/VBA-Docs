---
title: AppointmentItem.EndUTC property (Outlook)
keywords: vbaol11.chm3272
f1_keywords:
- vbaol11.chm3272
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.EndUTC
ms.assetid: c741e893-3a29-10cc-0730-a0796d8c2e4c
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.EndUTC property (Outlook)

Returns or sets a  **Date** value that represents the end date and time of the appointment expressed in the Coordinated Universal Time (UTC) standard. Read/write.


## Syntax

_expression_. `EndUTC`

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Remarks

Changing the value for the  **[AppointmentItem.End](Outlook.AppointmentItem.End.md)** property or the **[AppointmentItem.EndTimeZone](Outlook.AppointmentItem.EndTimeZone.md)** property will cause Outlook to recalculate the **EndUTC**.


## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]