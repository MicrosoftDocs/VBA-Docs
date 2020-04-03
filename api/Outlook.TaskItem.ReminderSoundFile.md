---
title: TaskItem.ReminderSoundFile property (Outlook)
keywords: vbaol11.chm1739
f1_keywords:
- vbaol11.chm1739
ms.prod: outlook
api_name:
- Outlook.TaskItem.ReminderSoundFile
ms.assetid: 29bfa689-08b6-f963-9ecb-3744b1032062
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.ReminderSoundFile property (Outlook)

Returns or sets a  **String** indicating the path and file name of the sound file to play when the reminder occurs for the Outlook item. Read/write.


## Syntax

_expression_. `ReminderSoundFile`

_expression_ A variable that represents a [TaskItem](Outlook.TaskItem.md) object.


## Remarks

This property is only valid if the  **[ReminderOverrideDefault](Outlook.TaskItem.ReminderOverrideDefault.md)** and **[ReminderPlaySound](Outlook.TaskItem.ReminderPlaySound.md)** properties are set to **True**.


## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]