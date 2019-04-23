---
title: AppointmentItem.ReminderSoundFile property (Outlook)
keywords: vbaol11.chm896
f1_keywords:
- vbaol11.chm896
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.ReminderSoundFile
ms.assetid: e3599e63-1300-7821-b94d-f8387a47e87d
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.ReminderSoundFile property (Outlook)

Returns or sets a  **String** indicating the path and file name of the sound file to play when the reminder occurs for the Outlook item. Read/write.


## Syntax

_expression_. `ReminderSoundFile`

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Remarks

This property is only valid if the  **[ReminderOverrideDefault](Outlook.AppointmentItem.ReminderOverrideDefault.md)** and **[ReminderPlaySound](Outlook.AppointmentItem.ReminderPlaySound.md)** properties are set to **True**.


## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]