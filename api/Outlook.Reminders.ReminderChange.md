---
title: Reminders.ReminderChange event (Outlook)
keywords: vbaol11.chm577
f1_keywords:
- vbaol11.chm577
ms.prod: outlook
api_name:
- Outlook.Reminders.ReminderChange
ms.assetid: 3af06d69-9a56-170e-9a51-c92d12efd293
ms.date: 06/08/2017
localization_priority: Normal
---


# Reminders.ReminderChange event (Outlook)

Occurs after a reminder has been modified.


## Syntax

_expression_. `ReminderChange`( `_ReminderObject_` )

_expression_ A variable that represents a [Reminders](Outlook.Reminders.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ReminderObject_|Required| **[Reminder](Outlook.Reminder.md)**|The  **Reminder** object that has been modified.|

## Example

The following Microsoft Visual Basic for Applications (VBA) example prompts the user with a message every time a reminder is modified.


```vb
Public WithEvents objReminders As Outlook.Reminders 
 
 
 
Sub Initialize_handler() 
 
 Set objReminders = Application.Reminders 
 
End Sub 
 
 
 
Private Sub objReminders_ReminderChange(ByVal ReminderObject As Reminder) 
 
 'Occurs when reminder is changed 
 
 MsgBox "The reminder " & ReminderObject.Caption & " has changed." 
 
End Sub
```


## See also


[Reminders Object](Outlook.Reminders.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]