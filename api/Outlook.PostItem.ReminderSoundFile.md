---
title: PostItem.ReminderSoundFile property (Outlook)
keywords: vbaol11.chm1579
f1_keywords:
- vbaol11.chm1579
ms.prod: outlook
api_name:
- Outlook.PostItem.ReminderSoundFile
ms.assetid: 9292a962-e7f9-75e0-20a0-716daf7d677f
ms.date: 06/08/2017
localization_priority: Normal
---


# PostItem.ReminderSoundFile property (Outlook)

Returns or sets a **String** indicating the path and file name of the sound file to play when the reminder occurs for the Outlook item. Read/write.


## Syntax

_expression_. `ReminderSoundFile`

_expression_ A variable that represents a [PostItem](Outlook.PostItem.md) object.


## Remarks

This property is only valid if the  **[ReminderOverrideDefault](Outlook.PostItem.ReminderOverrideDefault.md)** and **[ReminderPlaySound](Outlook.PostItem.ReminderPlaySound.md)** properties are set to **True**.


## See also


[PostItem Object](Outlook.PostItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]