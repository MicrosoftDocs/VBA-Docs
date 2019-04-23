---
title: AppointmentItem.StartUTC property (Outlook)
keywords: vbaol11.chm3271
f1_keywords:
- vbaol11.chm3271
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.StartUTC
ms.assetid: 8bfbf95f-bd88-acdc-f592-c41b454afe4b
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.StartUTC property (Outlook)

Returns or sets a  **Date** value that represents the start date and time of the appointment expressed in the Coordinated Universal Time (UTC) standard. Read/write.


## Syntax

_expression_. `StartUTC`

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Remarks

Changing the value for the  **[AppointmentItem.Start](Outlook.AppointmentItem.Start.md)** property or the **[AppointmentItem.StartTimeZone](Outlook.AppointmentItem.StartTimeZone.md)** property will cause Outlook to recalculate the value of **StartUTC**.


## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]