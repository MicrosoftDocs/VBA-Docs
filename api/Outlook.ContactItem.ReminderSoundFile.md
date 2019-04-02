---
title: ContactItem.ReminderSoundFile property (Outlook)
keywords: vbaol11.chm1107
f1_keywords:
- vbaol11.chm1107
ms.prod: outlook
api_name:
- Outlook.ContactItem.ReminderSoundFile
ms.assetid: aafbdc5b-816f-3605-d265-5da349e9e791
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.ReminderSoundFile property (Outlook)

Returns or sets a  **String** indicating the path and file name of the sound file to play when the reminder occurs for the Outlook item. Read/write.


## Syntax

_expression_. `ReminderSoundFile`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is only valid if the  **[ReminderOverrideDefault](Outlook.ContactItem.ReminderOverrideDefault.md)** and **[ReminderPlaySound](Outlook.ContactItem.ReminderPlaySound.md)** properties are set to **True**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]