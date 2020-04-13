---
title: Reminders.ReminderAdd event (Outlook)
keywords: vbaol11.chm576
f1_keywords:
- vbaol11.chm576
ms.prod: outlook
api_name:
- Outlook.Reminders.ReminderAdd
ms.assetid: cb1710f1-0c1d-eb71-e57f-6e33e3268576
ms.date: 06/08/2017
localization_priority: Normal
---


# Reminders.ReminderAdd event (Outlook)

Occurs after a reminder is added.


## Syntax

_expression_. `ReminderAdd`( `_ReminderObject_` )

_expression_ A variable that represents a [Reminders](Outlook.Reminders.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ReminderObject_|Required| **[_REMINDER]**|The **[Reminder](Outlook.Reminder.md)** object added to the collection.|

## Remarks

A reminder is not actually created until the associated Microsoft Outlook item has been saved. Therefore, this event will not occur until the associated item object has been saved.


## Example

The following example displays the date of the next reminder when a reminder is added to the collection.


```vb
Public WithEvents objReminders As Outlook.Reminders 
 
 
 
Sub Initialize_handler() 
 
 Set objReminders = Application.Reminders 
 
End Sub 
 
 
 
Private Sub objReminders_ReminderAdd(ByVal ReminderObject As Reminder) 
 
 'Occurs when a Reminder object is added to the collection using the user interface or object model 
 
 
 
 MsgBox "A new reminder is added that will fire at: " & _ 
 
 ReminderObject.NextReminderDate 
 
 
 
End Sub
```


## See also


[Reminders Object](Outlook.Reminders.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]