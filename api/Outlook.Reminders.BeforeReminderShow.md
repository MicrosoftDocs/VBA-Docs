---
title: Reminders.BeforeReminderShow event (Outlook)
keywords: vbaol11.chm575
f1_keywords:
- vbaol11.chm575
ms.prod: outlook
api_name:
- Outlook.Reminders.BeforeReminderShow
ms.assetid: 863859c0-a137-384d-80df-63fde038b533
ms.date: 06/08/2017
localization_priority: Normal
---


# Reminders.BeforeReminderShow event (Outlook)

Occurs before the  **Reminder** dialog box is displayed.


## Syntax

_expression_. `BeforeReminderShow`( `_Cancel_` )

_expression_ A variable that represents a [Reminders](Outlook.Reminders.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **True** to cancel the event. The default value is **False**.|

## See also


[Reminders Object](Outlook.Reminders.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]